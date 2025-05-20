import pyodbc
import logging
import os
import sys
import json
from mcp.server.fastmcp import FastMCP

# Configure basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='access_mcp.log',
    filemode='a'
)
logger = logging.getLogger("access_mcp")

# Create FastMCP server
mcp = FastMCP("Access DB")

# Default database path
default_db_path = r"M:\Quality System Database\old\BE\2025\Quality System Database_be 3-2-25 Post Compact.mdb"

# Database configuration
databases = {}

# Initialize databases from environment variable
try:
    # Check if environment variable exists for multiple databases
    db_config_str = os.environ.get("ACCESS_DB_CONFIG")
    if db_config_str:
        # Parse JSON configuration
        db_config = json.loads(db_config_str)
        for db_name, db_path in db_config.items():
            databases[db_name] = db_path
            logger.info(f"Added database '{db_name}' with path: {db_path}")
    
    # If no multiple database config found, use single database approach as fallback
    if not databases:
        single_db_path = os.environ.get("ACCESS_DB_PATH", default_db_path)
        default_db_name = "default"
        databases[default_db_name] = single_db_path
        logger.info(f"Using single database '{default_db_name}' with path: {single_db_path}")

except Exception as e:
    logger.error(f"Error parsing database configuration: {str(e)}")
    # Fallback to default
    databases["default"] = default_db_path
    logger.info(f"Falling back to default database with path: {default_db_path}")

def get_connection_string(db_path):
    """Create a connection string for the specified database path"""
    return (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
        r"ExtendedAnsiSQL=1;"
    )

def get_database_path(database_name):
    """Get the database path for the specified database name"""
    if database_name not in databases:
        available_dbs = ", ".join(databases.keys())
        raise ValueError(f"Database '{database_name}' not found. Available databases: {available_dbs}")
    return databases[database_name]

@mcp.tool()
async def list_databases() -> str:
    """List all available databases."""
    logger.info("Listing available databases")
    
    if not databases:
        return "No databases configured."
    
    results = ["Available Databases:", "-------------------"]
    for db_name, db_path in databases.items():
        exists = "✓" if os.path.exists(db_path) else "✗"
        results.append(f"{db_name}: {db_path} [{exists}]")
    
    return "\n".join(results)

@mcp.tool()
async def list_tables(database_name: str = "default") -> str:
    """List all tables in the specified Access database.
    
    Args:
        database_name: The name of the database to use (default: "default")
    """
    logger.info(f"Listing tables in database: {database_name}")
    
    try:
        db_path = get_database_path(database_name)
        
        # Check if database file exists
        if not os.path.exists(db_path):
            error_msg = f"Error: Database file not found at {db_path}. Please ensure the correct database file path is specified."
            logger.error(error_msg)
            return error_msg
            
        conn_str = get_connection_string(db_path)
        
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            tables = cursor.tables(tableType='TABLE')
            table_names = []
            
            for table_info in tables:
                table_name = table_info[2]
                if not table_name.startswith('MSys'):
                    table_names.append(table_name)
            
            if not table_names:
                return f"No tables found in database '{database_name}'."
                
            return f"Tables in database '{database_name}':\n" + "\n".join(table_names)
    except ValueError as e:
        return str(e)
    except pyodbc.Error as e:
        error_msg = f"Error connecting to database '{database_name}': {str(e)}"
        logger.error(error_msg)
        
        # Handle common ODBC driver errors
        if "IM002" in str(e):
            error_msg += "\n\nThe Microsoft Access ODBC driver is not installed or configured properly. "
            error_msg += "Please install the 'Microsoft Access Database Engine 2016 Redistributable' from Microsoft's website."
        
        return error_msg
    except Exception as e:
        logger.error(f"Error listing tables: {str(e)}")
        return f"Error listing tables: {str(e)}"

@mcp.tool()
async def query(sql: str, database_name: str = "default") -> str:
    """Execute a SQL query against the specified Access database.
    
    Args:
        sql: The SQL query to execute
        database_name: The name of the database to use (default: "default")
    """
    logger.info(f"Executing query in database '{database_name}': {sql}")
    
    try:
        db_path = get_database_path(database_name)
        
        # Check if database file exists
        if not os.path.exists(db_path):
            error_msg = f"Error: Database file not found at {db_path}. Please ensure the correct database file path is specified."
            logger.error(error_msg)
            return error_msg
        
        if not sql.strip().upper().startswith("SELECT"):
            return "Error: Only SELECT queries are allowed for security reasons."
        
        conn_str = get_connection_string(db_path)
        
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            
            # Check if cursor.description is None (no results)
            if cursor.description is None:
                return "Query executed successfully, but returned no results."
                
            columns = [column[0] for column in cursor.description]
            column_header = " | ".join(columns)
            separator = "-" * len(column_header)
            
            results = [f"Results from database '{database_name}':", column_header, separator]
            rows = cursor.fetchall()
            
            if not rows:
                results.append("No data found matching your query.")
                return "\n".join(results)
                
            for row in rows:
                formatted_values = []
                for value in row:
                    if value is None:
                        formatted_values.append("NULL")
                    else:
                        formatted_values.append(str(value))
                results.append(" | ".join(formatted_values))
                
            return "\n".join(results)
    except ValueError as e:
        return str(e)
    except pyodbc.Error as e:
        error_msg = f"Error executing query: {str(e)}"
        logger.error(error_msg)
        
        # Handle common ODBC driver errors
        if "IM002" in str(e):
            error_msg += "\n\nThe Microsoft Access ODBC driver is not installed or configured properly. "
            error_msg += "Please install the 'Microsoft Access Database Engine 2016 Redistributable' from Microsoft's website."
        elif "42S02" in str(e):
            error_msg += "\n\nThe table referenced in your query does not exist. Use the list_tables tool to see available tables."
            
        return error_msg
    except Exception as e:
        logger.error(f"Error executing query: {str(e)}")
        return f"Error executing query: {str(e)}"

@mcp.tool()
async def describe_table(table_name: str, database_name: str = "default") -> str:
    """Get the structure of a specific table including column names and types.
    
    Args:
        table_name: The name of the table to describe
        database_name: The name of the database to use (default: "default")
        
    Returns:
        The table structure as formatted text
    """
    logger.info(f"Describing table '{table_name}' in database '{database_name}'")
    
    try:
        db_path = get_database_path(database_name)
        
        # Check if database file exists
        if not os.path.exists(db_path):
            error_msg = f"Error: Database file not found at {db_path}. Please ensure the correct database file path is specified."
            logger.error(error_msg)
            return error_msg
        
        conn_str = get_connection_string(db_path)
            
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            
            # Get columns for the specified table
            columns = cursor.columns(table=table_name)
            
            # Format the results
            results = [f"Table: {table_name} (Database: {database_name})", "-" * (len(table_name) + len(database_name) + 20)]
            results.append("Column Name | Data Type | Nullable")
            results.append("-" * 50)
            
            column_list = []
            for column_info in columns:
                column_list.append(column_info)
                
            if not column_list:
                return f"Table '{table_name}' not found in database '{database_name}' or has no columns."
                
            for column_info in column_list:
                column_name = column_info[3]  # Column name is in the fourth position
                data_type = column_info[5]    # Data type is in the sixth position
                nullable = "Yes" if column_info[10] else "No"  # Nullable is in the 11th position
                
                results.append(f"{column_name} | {data_type} | {nullable}")
            
            return "\n".join(results)
    except ValueError as e:
        return str(e)
    except pyodbc.Error as e:
        error_msg = f"Error describing table: {str(e)}"
        logger.error(error_msg)
        
        # Handle common ODBC driver errors
        if "IM002" in str(e):
            error_msg += "\n\nThe Microsoft Access ODBC driver is not installed or configured properly. "
            error_msg += "Please install the 'Microsoft Access Database Engine 2016 Redistributable' from Microsoft's website."
            
        return error_msg
    except Exception as e:
        logger.error(f"Error describing table: {str(e)}")
        return f"Error describing table: {str(e)}"

@mcp.tool()
async def get_connection_info(database_name: str = "default") -> str:
    """Get information about the database connection configuration.
    
    Args:
        database_name: The name of the database to get info for (default: "default")
    """
    logger.info(f"Getting connection info for database '{database_name}'")
    
    try:
        db_path = get_database_path(database_name)
        
        # Check ODBC drivers
        available_drivers = []
        try:
            available_drivers = pyodbc.drivers()
        except Exception as e:
            logger.error(f"Error getting ODBC drivers: {str(e)}")
        
        # Check if database file exists
        db_exists = os.path.exists(db_path)
        
        info = [
            f"Database Connection Information for '{database_name}':",
            "-----------------------------------",
            f"Database path: {db_path}",
            f"Database file exists: {'Yes' if db_exists else 'No'}",
            "",
            "Available ODBC Drivers:",
            "-----------------------------------"
        ]
        
        if available_drivers:
            for driver in available_drivers:
                info.append(f"- {driver}")
        else:
            info.append("No ODBC drivers found.")
        
        # Check if Access driver is available
        access_driver_available = any("Access" in driver for driver in available_drivers)
        info.append("")
        info.append(f"Microsoft Access ODBC driver available: {'Yes' if access_driver_available else 'No'}")
        
        if not access_driver_available:
            info.append("")
            info.append("RECOMMENDATION:")
            info.append("To fix the missing Access driver, please install the 'Microsoft Access Database Engine 2016 Redistributable'")
            info.append("from Microsoft's website: https://www.microsoft.com/en-us/download/details.aspx?id=54920")
        
        return "\n".join(info)
    except ValueError as e:
        return str(e)
    except Exception as e:
        logger.error(f"Error getting connection info: {str(e)}")
        return f"Error getting connection info: {str(e)}"

if __name__ == "__main__":
    print("Starting Access Database MCP Server...")
    print(f"Available databases: {', '.join(databases.keys())}")
    
    # Print database paths and existence
    for db_name, db_path in databases.items():
        exists = os.path.exists(db_path)
        print(f"Database '{db_name}': {db_path} (exists: {exists})")
    
    # Architecture check
    if sys.maxsize > 2**32:
        print("WARNING: You are using 64-bit Python which may not be compatible with 32-bit Access drivers.")
    else:
        print("Using 32-bit Python which is compatible with Access database drivers.")
        
    print("Server running... Press Ctrl+C to exit")
    mcp.run()