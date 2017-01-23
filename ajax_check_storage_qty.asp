<%
response.expires=-1
response.expires=-1
  set conn=Server.CreateObject("ADODB.Connection")
  conn.Provider="Microsoft.ACE.OLEDB.12.0"
  conn.mode=adModeShareDenyNone or adModeRecursive or adModeReadWrite
  conn.Open(Server.Mappath("..\..\Desktop\每日进出库盘点计划打印-北京.accdb"))
  
  dim finalRet
  finalRet = True
  if request.servervariables("content_length") <= 0 then
    finalRet = False
  end if

  for each i in request.form
    dim ret
    rsq=request.form(i)
    taskID=i
    if len(rsq)>0 and len(taskID)>0 then
       sql="update 盘点记录 set RealStockQty=" & rsq & " where 序号=" & taskID
       conn.Execute sql,ret
       if ret<>1 then
          finalRet = False
       end if
    end if
  next
  if finalRet then
     response.write("True")
  end if
  conn.close
%>