<div align="center">

## NTService 1\.1


</div>

### Description

Microsoft's ActiveX Control NTSVC.OCX Version 1.1.

I've ported the Source-Code to VC++ 6, and updated it.

Added two methods and one property:

ReportStatus(), GetServiceStatusHandle() and WaitHint.

I think this is an "Open Source"-project and to publish

a new version, won't get me in a trouble regarding license regulations. I hope so.
 
### More Info
 
Short history: The reason why I updated this OCX is, I had to

write a Service-App for my company in VB. Which is a real

challenge. I've found this OCX and used it. But my App

has routines which need a long time during the Service-startup.

So the NT ServiceCM chancels my Service during Startup.

In the MSDN is mentioned a function called "SetServiceStatus()"

and a "WaitHint" property. With this, the NT SCM can be informed

about longer operations of the Service and it gives more time for Startup.

The Control uses this function internally, but has't put it out

to the interface. I changed the code, now it does. And it workes fine for

my purposes. I also updated the help-file.

Description for new interface (refer to the Help-file):

.GetServiceStatusHandle():

Returns the status handle for the Service.

It is a 32-bit integer value which is assigned to the Service during the StartService() operation.

It returns zero if the Service is not started.

.ReportStatus():

This function updates the service control manager's status information for the calling service.

.WaitHint:

Returns or sets an amount of extra time, in milliseconds,

which the NT Service Controller have to wait for the Service

at a start, stop, pause, or continue operation.


<span>             |<span>
---                |---
**Submitted On**   |2001-02-04 14:35:20
**By**             |[Harry Fischer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/harry-fischer.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD14561252001\.zip](https://github.com/Planet-Source-Code/harry-fischer-ntservice-1-1__1-15039/archive/master.zip)








