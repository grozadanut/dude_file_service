This is a Microsoft Visual Studio service for the DUDE driver for Datecs fiscal printers that extends the driver to read commands using a file on the file system containing commands that should be forwarded to the printer. Using this extension you could, for example, connect a Java app to the fiscal printer, DUDE being a .NET driver, thus not being able to use it in Java.

# Usage
1. Install the service using InstallUtil utility by issuing `installutil DUDEFileService.exe` command provided by the Developer Command Prompt for VS. More info here: https://learn.microsoft.com/en-us/dotnet/framework/windows-services/how-to-install-and-uninstall-services
2. Set **ecr_folder** environment variable pointing to the path where you want to read command files from.
3. Copy command files to ecr_folder location. The file should have **.in** extension
4. The result of the command will be written in a file with the same name, appending **_result** suffix

Note: currently the only implemented communication protocol is through LAN.

# Command file structure
## Execute script
line 1: ecr_ip<br/>
line 2: ecr_port<br/>
other lines: commands(as specified in the DUDE driver specification)

## Download the MF report
`raportmf&{0}&{1}&{2}` <br/>
{0} = startDateTime with format: dd-MM-yy HH:mm:ss 'DST'(if start time is in DST); example: 01-01-24 10:00:00 DST<br/>
{1} = endDateTime same format as start date<br/>
{2} = chosenDirectory to export the report to

## Read receipts from ECR
`receipts&{0}&{1}` <br/>
{0} = startDateTime with format: dd-MM-yy HH:mm:ss 'DST'(if start time is in DST); example: 01-01-24 10:00:00 DST<br/>
{1} = endDateTime same format as start date<br/>

The results will be saved to the same location of the input file appending `_result` to the filename.