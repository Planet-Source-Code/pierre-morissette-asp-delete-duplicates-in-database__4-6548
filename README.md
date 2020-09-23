<div align="center">

## Asp delete duplicates in Database


</div>

### Description

This code delete duplicate records in MSAcces database.
 
### More Info
 
The user must define 4 variables (lines 29 to 32) in order to configure the code for his/her personnal use.

The author is not responsible for any lost of your data.

Read the code comments before using it.

Works only with dsn-less connexion.

This code returns the number of record in the database and the number of deleted duplicate records.

The author suggest you to use the code first on a copy of your database to prevent unwanted deletion.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pierre Morissette](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pierre-morissette.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pierre-morissette-asp-delete-duplicates-in-database__4-6548/archive/master.zip)

### API Declarations

Feel free to use the code as long as you leave the author name and email adress.


### Source Code

```
<%@ LANGUAGE = "VBScript" %>
<% option explicit %>
<!-- IMPORTANT....set server,scripttimeout to any value in seconds
If your database is very large, it will take a long time to execute -->
<%server.scripttimeout=600%>
<!--
IMPORTANT!!
Use this script at your own risk.
The author is not responsible for any lost data.
Try this script on a backup database first to see the results.
This script finds and eliminate duplicate records (duplicate of "url" field content) in table "mes_sites" of database "../../db/signets.mdb"
It finds the value of the field named "ID" (primary key field) and use this value to delete duplicate records.
If you want to use this script, you will have to change those value in this page :
the name of the field to look for duplicate value (dim myfield)
the name of the database (dim mydatabase)
IMPORTANT
This code uses a dsn-less connection to the database
the name of the table (dim mytable)
the name of your primary key field (dim myprimarykey)
It is probably not the best way to to the job, but it works.
If you know a better way, please contact me.
Thanks
Author : Pierre Morissette
mail: pierre@hawk.igs.net
 -->
<%
'IMPORTANT user must define those variables CAREFULLY
Dim mydatabase,mytable,myprimarykey,myfield
mydatabase="../../db/signets.mdb"
mytable="mes_sites"
myprimarykey="id"
myfield="url"
dim SQL,conn,rs,nb,i,nbtot,valurl,nbdup,nbdup2,valret,validdup,arr,nbarr
Set Conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Conn.Open "DBQ=" & Server.Mappath(mydatabase) & ";Driver={Microsoft Access Driver (*.mdb)};"
%>
<%
sql="select count(*) as nb from "
sql=sql & mytable
Set RS = conn.Execute(SQL)
'calcul du nombre de fiches, count the number of records in table
arr=""
nbtot=rs("nb")
nbtot=cint(nbtot)
response.write nbtot
response.write " fiches / files"
response.write "<hr>"
sql="select "
sql = sql & myfield
sql=sql & ","
sql=sql & myprimarykey
sql=sql & " from "
sql=sql & mytable
sql=sql & " order by "
sql=sql & myfield
set rs=conn.execute(sql)
'selectionner la valeur de myfield(i), select the value of field myfield # i
for i=0 to (nbtot- 1)
rs.movefirst
rs.move(i)
valurl= rs.fields(myfield)
' vérifier si valeur dupliquée, check if the value of the field is a duplicate value
if valurl=valret then
validdup= rs.fields(myprimarykey)
' ajouter l'id de la fiche à la liste des duplicats, add id value to the array if duplicate value
arr= arr & validdup
arr = arr & ","
else
end if
'remind the last value to compare to next one
valret = valurl
next
rs.close
set rs=nothing
'écrire la liste des duplicats, writes the list of all id value that contains duplicates
response.write "Records that contains duplicate data in field myfield"
response.write "<br>"
'now use the array created to delete records
'create array
if arr = "" then
response.write "There is no duplicate record."
else
arr=left(arr,len(arr)-1)
response.write arr
arr=split(arr,",",-1,1)
nbarr = ubound(arr)
nbarr=nbarr + 1
response.write "<br>"
response.write "Number of duplicate records :"
response.write nbarr
response.write "<br>"
'number of records to duplicate
for i=0 to nbarr-1
SQL = "delete from "
sql=sql & mytable
SQL = SQL & " WHERE "
sql=sql & myprimarykey
sql = sql & " ="
sql=sql & arr(i)
Set RS = conn.Execute(SQL)
next
Response.write "All the duplicate records are now deleted."
end if
set rs=nothing
conn.close
set conn = nothing
 %>
```

