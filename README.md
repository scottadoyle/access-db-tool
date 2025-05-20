# Access Database MCP Server

This is a simple Model Context Protocol (MCP) server that connects to a local Microsoft Access database for use with Claude Desktop.

## Prerequisites

- **Python 3.10 or higher (32-bit version recommended)**
- Microsoft Access Database
- The Microsoft Access ODBC driver installed (Microsoft Access Database Engine 2016 Redistributable)
- Model Context Protocol (MCP) Python SDK

## Important: 32-bit vs 64-bit Compatibility

**The Microsoft Access ODBC driver is usually 32-bit, which requires a 32-bit Python installation.** 

If you have both 32-bit and 64-bit Python installed:
- Check your Python installations with `py -0` in Command Prompt
- Use the 32-bit version explicitly with `py -3.12-32` (adjust version number as needed)
- Ensure you install all required packages in the 32-bit Python environment

## Setup

1. Identify which Python installations you have:
```
py -0
```

2. Install the required packages in your 32-bit Python environment:
```
py -3.12-32 -m pip install pyodbc
py -3.12-32 -m pip install mcp
```

   If `pip install mcp` fails, install directly from GitHub:
```
py -3.12-32 -m pip install git+https://github.com/modelcontextprotocol/python-sdk.git
```

3. The database path can now be configured in two ways:
   - **Recommended:** Set the `ACCESS_DB_PATH` environment variable in the Claude Desktop configuration file (as shown in the Configuration section below)
   - As a fallback, you can also modify the `default_db_path` in the `access_simple.py` script

4. Run the server using 32-bit Python:
```
py -3.12-32 access_simple.py
```

## Configuring Claude Desktop

1. Locate the Claude Desktop configuration file:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - Mac: `~/Library/Application Support/Claude/claude_desktop_config.json`

2. Add the following configuration, specifying your 32-bit Python and database path:
```json
{
  "mcpServers": {
    "access_db": {
      "command": "py",
      "args": [
        "-3.12-32",
        "C:/Users/sdoyle/Desktop/Projects/Quality System Database AI/access_simple.py"
      ],
      "env": {
        "ACCESS_DB_PATH": "M:/Quality System Database/old/BE/2025/Your_Database_File.mdb"
      }
    }
  }
}
```

3. Adjust the paths and Python version as needed for your environment
4. Save the file and restart Claude Desktop

## Available Tools

The MCP server provides the following tools:

- `list_tables` - Get a list of all tables in the database
- `describe_table` - Get the structure of a specific table including column names and types
- `query` - Execute a SQL query against the database (SELECT only for security)
- `get_connection_info` - Get information about the database connection configuration

## Example Prompts for Claude

Here are some example prompts you can use with Claude:

```
Use the access_db MCP server to list all tables in my Access database.

Use the access_db MCP server to describe the structure of the tblFPRawData table.

Use the access_db MCP server to get the connection information and check if my configuration is correct.

Use the access_db MCP server to run a SQL query that selects the first 10 records from the tblFPRawData table.
```

## Troubleshooting

If you encounter connection issues, try the following:

1. Verify you're using a 32-bit Python environment:
   - Check with `py -0` to see all installed Python versions
   - Make sure you're using the 32-bit version with the `-3.12-32` flag (adjust version as needed)

2. Check if the MCP package is installed correctly:
   - Try reinstalling it with `py -3.12-32 -m pip install mcp`
   - Or install from GitHub: `py -3.12-32 -m pip install git+https://github.com/modelcontextprotocol/python-sdk.git`

3. Use the `get_connection_info` tool to check your database path and ODBC driver availability

4. Make sure the Microsoft Access ODBC driver is installed (Microsoft Access Database Engine 2016 Redistributable)
   - Specifically, install the 32-bit version even if you're on a 64-bit Windows system

5. Check that the database file path is correct and accessible

6. Look at the MCP server logs for more detailed error information:
   - Windows: `%APPDATA%\Claude\logs\mcp*.log`

7. Try running the server manually from the command line before starting Claude Desktop:
   ```
   py -3.12-32 access_simple.py
   ```
   Keep this window open and start Claude Desktop in a new window
