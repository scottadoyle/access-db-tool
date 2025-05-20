import pyodbc
import logging
import os
import sys
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

# Path to the Access database file
default_db_path = r"M:\Quality System Database\old\BE\2025\Quality System Database_be 3-2-25 Post Compact.mdb"

# Get database path from environment variable if set, otherwise use default
db_path = os.environ.get("ACCESS_DB_PATH", default_db_path)
logger.info(f"Using database path: {db_path}")

# Access Database connection string
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={db_path};"
    r"ExtendedAnsiSQL=1;"
)

@mcp.tool()
async def list_tables() -> str:
    """List all tables in the Access database."""
    logger.info("Listing tables")
    
    # Check if database file exists
    if not os.path.exists(db_path):
        error_msg = f"Error: Database file not found at {db_path}. Please ensure the correct database file path is specified."
        logger.error(error_msg)
        return error_msg
        
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            tables = cursor.tables(tableType='TABLE')
            table_names = []
            
            for table_info in tables:
                table_name = table_info[2]
                if not table_name.startswith('MSys'):
                    table_names.append(table_name)
            
            if not table_names:
                return "No tables found in the database."
                
            return "\n".join(table_names)
    except pyodbc.Error as e:
        error_msg = f"Error connecting to database: {str(e)}"
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
async def query(sql: str) -> str:
    """Execute a SQL query against the Access database."""
    logger.info(f"Executing query: {sql}")
    
    # Check if database file exists
    if not os.path.exists(db_path):
        error_msg = f"Error: Database file not found at {db_path}. Please ensure the correct database file path is specified."
        logger.error(error_msg)
        return error_msg
    
    if not sql.strip().upper().startswith("SELECT"):
        return "Error: Only SELECT queries are allowed for security reasons."
    
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            
            # Check if cursor.description is None (no results)
            if cursor.description is None:
                return "Query executed successfully, but returned no results."
                
            columns = [column[0] for column in cursor.description]
            column_header = " | ".join(columns)
            separator = "-" * len(column_header)
            
            results = [column_header, separator]
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
async def describe_table(table_name: str) -> str:
    """Get the structure of a specific table including column names and types.
    
    Args:
        table_name: The name of the table to describe
        
    Returns:
        The table structure as formatted text
    """
    logger.info(f"Describing table: {table_name}")
    
    # Check if database file exists
    if not os.path.exists(db_path):
        error_msg = f"Error: Database file not found at {db_path}. Please ensure the correct database file path is specified."
        logger.error(error_msg)
        return error_msg
        
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            
            # Get columns for the specified table
            columns = cursor.columns(table=table_name)
            
            # Format the results
            results = [f"Table: {table_name}", "-" * (len(table_name) + 7)]
            results.append("Column Name | Data Type | Nullable")
            results.append("-" * 50)
            
            column_list = []
            for column_info in columns:
                column_list.append(column_info)
                
            if not column_list:
                return f"Table '{table_name}' not found or has no columns."
                
            for column_info in column_list:
                column_name = column_info[3]  # Column name is in the fourth position
                data_type = column_info[5]    # Data type is in the sixth position
                nullable = "Yes" if column_info[10] else "No"  # Nullable is in the 11th position
                
                results.append(f"{column_name} | {data_type} | {nullable}")
            
            return "\n".join(results)
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
async def get_connection_info() -> str:
    """Get information about the database connection configuration."""
    logger.info("Getting connection info")
    
    # Check ODBC drivers
    available_drivers = []
    try:
        available_drivers = pyodbc.drivers()
    except Exception as e:
        logger.error(f"Error getting ODBC drivers: {str(e)}")
    
    # Check if database file exists
    db_exists = os.path.exists(db_path)
    
    info = [
        "Database Connection Information:",
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

if __name__ == "__main__":
    print("Starting Access Database MCP Server...")
    print(f"Database path: {db_path}")
    print(f"Database exists: {os.path.exists(db_path)}")
    
    # Architecture check
    if sys.maxsize > 2**32:
        print("WARNING: You are using 64-bit Python which may not be compatible with 32-bit Access drivers.")
    else:
        print("Using 32-bit Python which is compatible with Access database drivers.")
        
    print("Server running... Press Ctrl+C to exit")
    mcp.run()