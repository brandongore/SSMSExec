# SSMSExec
SSMSExec is an SSMS extension which takes the current query and runs executables against it with configurable parameters based on options and updates the query in the window.

## Setup and Running
- Open options -> SSMSExec -> General
- Set the ExeLocation parameter to the full path of the exe you want to run, eg. c:\dev\python.exe
- Set the Exe Parameters, eg. Exe Parameter 1 ( -m sqlfluff fix - --dialect=tsql)
- Open a new Query window and input a query.
- Open the tools window and click the "Update Query" (or configured button name) to run the query as standard input to the exe with the parameters

Possible uses:
- Format the current query using sqlFluff formatter  ( python -m sqlfluff fix - --dialect=tsql )
- Templating, build a rust exe using sqlparser library to parse the query into an Abstract Syntax Tree (AST) and replace special characters or blocks.