<div align="center">

## WebSvc


</div>

### Description

WebSvc is a VB6 Client that connects to a web Service set up for Basic Authentication, using MSSoap30.

In this example program, WebSvc is used to continually submit files to a web service.

WHen started, WebSvc waits for a file matching a file mask to show up in a designated folder, then sends that file a record at a time to a web service, logging activity in a continualy appended log file, renaming the transmitted file for autoarchiving, then idles until another file shows up.

MS examples for MSSOAP30 using basic authentication web service did not work. Solution is include login/password in the connection string, and to implement iHeaderHandler event to handle the authentication challenge that comes back on the serivice function call.
 
### More Info
 
All parameters are customer entered, and saved in a WebSvc.Set file in the app folder.

User will need to mdify the service call(s) to those defined on their target Web Service. Some basic knowledge re how a web service works is assumed. This project uses the MSSOAP30.DLL, and the MSXML3.dll

A log file is appended to showing all file transmissions, connections, start stops of the process (WebSvc.Log)

Process stops when a connection fails, or a transmission attempt fails. That event is logged. User must at that time ascertain why the web servcice is not available, then restart once the web service is running/available again.


<span>             |<span>
---                |---
**Submitted On**   |2007-12-14 09:17:44
**By**             |[Doug V\. Taylor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/doug-v-taylor.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[WebSvc20938112142007\.zip](https://github.com/Planet-Source-Code/doug-v-taylor-websvc__1-69760/archive/master.zip)








