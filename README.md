# Remote Shell Command Service

Remote Shell Command Service is a Windows Service that makes use of web interaction to retrieve remotely submitted DOS commands. By periodically polling a config file controlled web address, Remote Shell Command Service checks for DOS commands it has not run yet. If it comes across an unexecuted command ID, it updates the database sitting on the web server, indicating by means of its MAC address that it has now responded to that specific command. This service is particularly useful when managing a large LAN. For example, to automate the resetting of all computers on the network using a 'shutdown -r' DOS call.

Created by Craig Lotter, December 2005

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2003
Implements concepts such as Service Programming and Shell Scripting.
Level of Complexity: simple
