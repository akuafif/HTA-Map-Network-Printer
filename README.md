# HTA-Map-Network-Printer
2 HTAs for Login to Print Server and Mapping Network Printer.  
If the print server is already logged in, you can just use the `Map_Printer.hta`.

# Screenshots
![Login](https://cdn.discordapp.com/attachments/240938506103422976/780824595217252382/unknown.png)
![Map](https://cdn.discordapp.com/attachments/240938506103422976/780824779091476510/unknown.png)

# How is it being done?
Using [ProcessOutputMonitor](https://github.com/ru-rararu-ra-rurararu-ra/ProcessOutputMonitor) Class.  
`ProcessOutputMonitor` makes it possible to run multiple command at once.   
It will spawn a child process and monitor its output.  
This will not make HTA GUI hangs and shows not responding in the title bar if the code is running long command.  

`net use` for logging in.    
`net view` for fetching printers from print server. `net view` is way faster than `WMIService.ExecQuery` in vbs.  
`rundll32 printui.dll,PrintUIEntry` for printer ultilities.

# Config File
Edit the config.xml file to be used in your environment.  
Run `ipconfig` in cmd to know what's your intranet dns suffix.  
You can add or delete Print Servers and it's prefix.

```
<?xml version="1.0" encoding="UTF-8" ?>
<CONFIGURATION>
	<INTRANETDNSSUFFIX>domain.com</INTRANETDNSSUFFIX>
	<SERVER>
		<HOSTNAME>Print Server 1</HOSTNAME>
		<PREFIX>Printer Prefix 1</PREFIX>
	</SERVER>
	<SERVER>
		<HOSTNAME>Print Server 2</HOSTNAME>
		<PREFIX>Printer Prefix 2 </PREFIX>
	</SERVER>
	<SERVER>
		<HOSTNAME>Print Server 3</HOSTNAME>
		<PREFIX>Printer Prefix 3</PREFIX>
	</SERVER>
</CONFIGURATION>
```


