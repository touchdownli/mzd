<%
sub showTable(response, rs)
response.write("<table>")
<!--print td name-->
if not rs.EOF then
    response.write("<tr>")
    for each x in rs.Fields
        response.write("<td><b>" & x.name & "</b></td>")
    next
    response.write("</tr>")
end if

do until rs.EOF
  response.write("<tr>")
  for each x in rs.Fields
    response.write("<td>" & x.value & "</td>")
  next
  response.write("</tr>")
  rs.MoveNext
loop
response.write("</table>")
end sub
%>

<%
sub conAndShow()
  response.expires=-1
  sql="SELECT * FROM 盘点记录 where 盘点日期=#2014/5/20#"
  <!--sql=sql & request.querystring("q") & "#"-->

  set conn=Server.CreateObject("ADODB.Connection")
  conn.Provider="Microsoft.ACE.OLEDB.12.0"
  conn.mode=adModeShareDenyNone or adModeRecursive
  conn.Open(Server.Mappath("..\..\Desktop\每日进出库盘点计划打印-北京.accdb"))
  
  set rs=Server.CreateObject("ADODB.recordset")
  rs.Open sql,conn
  
  call showTable(response, rs)
  conn.Close
end sub
%>
<%
sub PrintErr()
            set Err=Server.GetLastError()
            Response.Write "<ul>"
            Response.Write "<li>错误代码：0x" & Hex(Err.Number) & "</li>"
            Response.Write "<li>错误描述：" & Err.Description & "</li>"
            Response.Write "<li>错误来源：" & Err.Source & "</li>"
            Response.Write "</ul>"
end sub

function CheckStaffNo(inStr)
  set regEx = new RegExp
  regEx.Pattern = "^[0-9a-z_]+$"
  regEx.IgnoreCase = True
  regEx.Global = True
  CheckStaffNo = regEx.Test(inStr)
end function

sub alert(msg)
  alertScript = "<script language='javascript'>alert('" & msg & "')</script>"
  response.write(alertScript)
end sub
%>

<!--#include file="JSON_2.0.4.asp"-->
<%
function makePagesJSONStr(rs, isTakeNewTask, staffNo)
  redim ret(0)
  makePagesJSONStr=ret
  pageCnt=0
  taskStaffCnt=0
  taskCnt=0
  if isTakeNewTask then
    do until rs.EOF
      pageCnt = pageCnt + 1
      if taskStaffCnt = 0 or taskStaffCnt = rs.Fields.Item("TaskStaffCnt").value then
        taskStaffCnt = rs.Fields.Item("TaskStaffCnt").value
      elseif taskStaffCnt <> rs.Fields.Item("TaskStaffCnt").value then
        response.write("taskStaffCnt is not same for each task")
        exit function
      end if
      taskStaffCnt=rs.Fields.Item("TaskStaffCnt").value
      rs.MoveNext
    loop
    
    if isnull(taskStaffCnt) then
      taskStaffCnt = 1
    end if
    if pageCnt < 1 or taskStaffCnt < 1 then
      alert("pageCnt < 1 or taskStaffCnt < 1")
      exit function
    end if
    
    'shang qu zheng
    taskCnt=Int(pageCnt / taskStaffCnt + 0.5)
  else
    do until rs.EOF
      taskCnt = taskCnt + 1
      rs.MoveNext
    loop
  end if

  redim pages(taskCnt-1)
  redim ret(taskCnt-1)

  i=0
  rs.MoveFirst
  do until rs.EOF or i >= taskCnt
    dim tmpDict
    set tmpDict=jsObject()
    tmpDict("PartsCode")=rs.Fields.Item("PartsCode").value
    tmpDict("DescCh")=rs.Fields.Item("DescCh").value
    tmpDict("StorageCode")=rs.Fields.Item("StorageCode").value
    tmpDict("StockQty")=rs.Fields.Item("StockQty").value
    'tmpDict("RealStockQty")=rs.Fields.Item("RealStockQty").value
    tmpDict("序号")=rs.Fields.Item("序号").value
    tmpDict("CheckedTimes")=rs.Fields.Item("CheckedTimes").value
    set pages(i)=tmpDict
    
    ret(i) = rs.Fields.Item("序号").value
    
    i = i + 1
    rs.MoveNext
  loop
  
  dim pagesJSONStr
  pagesJSONStr=toJSON(pages)
  dim jsStrTpl
  jsStrTpl="<script language='javascript'>var g_staffNo=STAFFNO;var g_pages=PAGES_DATA_VECTOR</script>"
  dim jsStr
  jsStr=replace(jsStrTpl, "PAGES_DATA_VECTOR", pagesJSONStr)
  jsStr=replace(jsStr, "STAFFNO", staffNo)
  
  response.write(jsStr)
  
  makePagesJSONStr = ret
end function

function makeTask(staffNo)
  makeTask=False
  if isempty(staffNo) or isnull(staffNo) then 
    exit function
  end if
  if not CheckStaffNo(staffNo) then
    alert("staffNo is invalid")
    exit function
  end if
  
  set conn=Server.CreateObject("ADODB.Connection")
  conn.Provider="Microsoft.ACE.OLEDB.12.0"
  conn.mode=adModeShareDenyNone or adModeRecursive or adModeReadWrite
  conn.Open(Server.Mappath("..\..\Desktop\每日进出库盘点计划打印-北京.accdb"))
  
  set rs=Server.CreateObject("ADODB.recordset")

  sql="SELECT * FROM 盘点记录 where RealStockQty=-1 and StaffNo="
  sql = sql & staffNo & " order by StorageCode"
  'sql="SELECT * FROM 盘点记录 where RealStockQty=-1 and 盘点日期=#2014/5/20#"
  'alert(sql)
  rs.Open sql,conn
  
  dim isTakeNewTask
  isTakeNewTask = False
  if rs.EOF then
    isTakeNewTask = True
    rs.close
    sql="SELECT * FROM 盘点记录 where RealStockQty=-1 and StaffNo=-1 order by StorageCode"
    rs.Open sql,conn
  end if

  ret=makePagesJSONStr(rs, isTakeNewTask, staffNo)
  
  if isTakeNewTask and ubound(ret) > 0 then
    taskIDsStr = join(ret,",")
    sql="update 盘点记录 set TaskStaffCnt=TaskStaffCnt-1 where StaffNo=-1"
    conn.Execute sql
    sql = "update 盘点记录 set StaffNo="& staffNo & " where 序号 in (" & taskIDsStr & ")"
    'alert(sql)
    'update staffNo
    conn.Execute sql
  end if
  
  if ubound(ret) > 0 or not isnull(ret(0)) then
    makeTask=True
  end if
  
  conn.close
end function
%>

<!DOCTYPE html>
<html>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<body onload='onBodyLoad()'>

<%
if Request.QueryString = "" or not makeTask(Request.QueryString("StaffNo")) then
%>
 <span id='date'></span><br/>
 <input id='takeTaskBtn' type='button' value='领取任务' onclick='onTakeTaskBtnClick()'/>
 工号:<input id='StaffNoText' style='text' size='3' /><br/>
<%
end if
%>
 <table id='pageTable'>
  <tbody>
    <tr><td>任务序号：</td><td id='序号'></td></tr>
    <tr><td>零 件 号：</td><td id='PartsCode' style="color: red;"></td></tr>
    <tr><td>零件描述：</td><td id='DescCh'></td></tr>
    <tr><td>储 位 号：</td><td id='StorageCode' style="color: red;"></td></tr>
    <tr><td>实际库存：</td><td><input style='text' id='RealStockQty' value='0' size='6'></input></td></tr>
    <tr>
      <td><input id='CheckAllBtn' type='button' value='最终上传' onclick='onCheckAllBtnClick()'/></td>
      <td><input id='CheckBtn' type='button' value='确 认' onclick='onCheckBtnClick()'/></td>
    </tr>
    <tr><td style="height: 70px;"></td></tr>
    <tr>
      <td><input id='preBtn' type='button' onclick='onClickPageBtn(-1)' value='上一页'></td>
      <td><input id='nextBtn' type='button' onclick='onClickPageBtn(1)' value='下一页'></td>
    </tr>
    <tr>
      <td><input id='firstBtn' type='button' onclick='onClickPageBtn(0)' value='第一页'></td>
      <td><input id='lastBtn' type='button' onclick='onClickPageBtn(2)' value='最后一页'></td>
    </tr>
    <tr>  
      <td>第<span id="spanPageNum"></span>页</td>
      <td>共<span id="spanTotalPage"></span>页</td>
    </tr>
  </tbody>
 </table>
<script language="javascript">
  var g_curIndex = 0;
  var g_totalCnt = 0;
  if (!(typeof(g_pages)==='undefined'))
  {
    g_totalCnt = g_pages.length;
  }
  //pageInc 0,move to first;2,move to last;
  function onClickPageBtn(pageInc)
  {
    var lastIndex = window.g_curIndex;
    if (pageInc == 0)
    {
      window.g_curIndex = 0;
    }
    else if(pageInc == 2)
    {
      window.g_curIndex = g_totalCnt - 1;
    }
    else if(pageInc == 1 || pageInc == -1)
    {
      window.g_curIndex += pageInc;
    }
    else
    {
      alert("pageInc is " + pageInc);
      return;
    }
    
    if (window.g_curIndex < 0)
    {
      //alert("window.g_curIndex is " + window.g_curIndex);
      window.g_curIndex = 0;
    }
    if (window.g_curIndex >= g_totalCnt)
    {
      window.g_curIndex = g_totalCnt - 1;
    }
    saveRSQ(lastIndex);
    showPage(window.g_curIndex)
  }
  
  function saveRSQ(pageIndex)
  {
    if (pageIndex >= g_totalCnt)
    {
      return;
    }
    //save realstockqty
    var rsq = parseInt(document.getElementById("RealStockQty").value);
    g_pages[pageIndex]["RealStockQty"] = rsq >= 0 ? rsq : ""; 
  }
  
  function showPage(pageIndex)
  {
    if (window.g_curIndex == 0)
    {
      //set disable first
      document.getElementById("firstBtn").disabled=true
      document.getElementById("preBtn").disabled=true
      document.getElementById("lastBtn").disabled=false
      document.getElementById("nextBtn").disabled=false
    }
    if (window.g_curIndex == g_totalCnt - 1)
    {
      //set disable last
      document.getElementById("lastBtn").disabled=true
      document.getElementById("nextBtn").disabled=true
      document.getElementById("firstBtn").disabled=false
      document.getElementById("preBtn").disabled=false
    }
    
    if (window.g_curIndex > 0 && window.g_curIndex < g_totalCnt-1)
    {
      document.getElementById("lastBtn").disabled=false
      document.getElementById("nextBtn").disabled=false
      document.getElementById("firstBtn").disabled=false
      document.getElementById("preBtn").disabled=false
    }
    
    document.getElementById("StorageCode").innerHTML = g_pages[pageIndex]["StorageCode"];
    document.getElementById("spanPageNum").innerHTML = window.g_curIndex + 1;
    document.getElementById("spanTotalPage").innerHTML = g_totalCnt;
    
    document.getElementById("PartsCode").innerHTML = g_pages[pageIndex]["PartsCode"];
    document.getElementById("DescCh").innerHTML = g_pages[pageIndex]["DescCh"];
    document.getElementById("序号").innerHTML = g_pages[pageIndex]["序号"];
    
    var RealStockQty = document.getElementById("RealStockQty");
    if (typeof(g_pages[pageIndex]["RealStockQty"]) === 'undefined')
    {
      RealStockQty.value = "";
    }
    else
    {
      RealStockQty.value = g_pages[pageIndex]["RealStockQty"];
    }
    
    if (g_pages[pageIndex]["IsMatch"] == "True" || g_pages[pageIndex]["CheckedTimes"] > 2)
    {
      //alert("IsMatch or CheckedTimes > 2");
      document.getElementById("CheckBtn").disabled = true;
      document.getElementById("RealStockQty").disabled = true;
    }
    else
    {
      document.getElementById("CheckBtn").disabled = false;
      document.getElementById("RealStockQty").disabled = false;
    }
  }
  
  function onBodyLoad()
  {
    if (typeof(g_pages) === 'undefined')
    {
      document.getElementById("pageTable").style.visibility="hidden";
    }
    else
    {
      recoverTaskStatus();
      showPage(0);
      document.getElementById("CheckAllBtn").disabled = true;
    }
  }
  
  function onTakeTaskBtnClick()
  {
    window.location.replace("Default.asp?StaffNo=" + document.getElementById("StaffNoText").value);
  }
  
  function onCheckBtnClick()
  {
    saveRSQ(window.g_curIndex);
    var rsq = document.getElementById("RealStockQty").value;
    if (rsq.length==0)
    {
      alert("不能为空");
      return;
    }
    
    savePages2Cookie();
    
    if (rsq == g_pages[window.g_curIndex]["StockQty"])
    {
      g_pages[window.g_curIndex]["IsMatch"] = "True";
      document.getElementById("CheckBtn").disabled = true;
      document.getElementById("RealStockQty").disabled = true;
      onClickPageBtn(1);
    }
    else
    {
      alert("数据不一致，请重新检查");
    }
    g_pages[window.g_curIndex]["CheckedTimes"] += 1;
    if (g_pages[window.g_curIndex]["CheckedTimes"] > 2)
    {
      alert("是否确认最终数量？");
      document.getElementById("CheckBtn").disabled = true;
      document.getElementById("RealStockQty").disabled = true;
      onClickPageBtn(1);
    }
    var allSure = true;
    for (var i=0;i<g_totalCnt;i++)
    {
      if (g_pages[i]["IsMatch"] == "True" || g_pages[i]["CheckedTimes"] > 2)
      {
        continue;
      }
      allSure = false; 
    }
    if (allSure)
    {
      document.getElementById("CheckAllBtn").disabled = false;
    }

  }
  
  function savePages2Cookie()
  {
    var pagesJSON = {};
    for (var i=0; i<g_totalCnt;++i)
    {
      pagesJSON[window.g_pages[i]["序号"]] = {};
    }
    for (var i=0; i<g_totalCnt;++i)
    {
      pagesJSON[window.g_pages[i]["序号"]]["RealStockQty"]=window.g_pages[i]["RealStockQty"];
      pagesJSON[window.g_pages[i]["序号"]]["PartsCode"]=window.g_pages[i]["PartsCode"];
    }
    var val = json2str(pagesJSON);
    setCookie(window.g_staffNo, val, 20);
  }
  
  function onCheckAllBtnClick()
  {
    var xmlhttp;
    if (window.XMLHttpRequest)
    {// code for IE7+, Firefox, Chrome, Opera, Safari
      xmlhttp=new XMLHttpRequest();
    }
    else
    {// code for IE6, IE5
      xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
    }
    
    xmlhttp.onreadystatechange=function()
    {
      if (xmlhttp.readyState == 4)
      {
          if (xmlhttp.status != 200)
          {
            alert("数据上传失败,http status:" + xmlhttp.status);
            return;
          }
          
          if (xmlhttp.responseText == "True")
          {
            deleteCookie(window.g_staffNo);
            alert("数据上传成功！");
          }
          else
          {
          	alert("SQL执行失败,请重新上传！");
          }
      }
    }
    
    var postStr = "";
    for (var i=0;i<window.g_totalCnt;i++)
    {
        var rsq = g_pages[i]["RealStockQty"];
        var taskID = g_pages[i]["序号"];
        if (typeof(rsq) === 'undefined' || rsq.length==0)
        {
          alert("第" + (i+1) + "页不能为空");
          return;
        }
        postStr += taskID + "=" +rsq + "&";
    }
    xmlhttp.open("POST","ajax_check_storage_qty.asp", false);
    xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    xmlhttp.send(postStr);
  }
  
  function onOldCheckBtnClick()
  {
      saveRSQ(window.g_curIndex);
      var xmlhttp;
      var rsq = document.getElementById("RealStockQty").value;
      var taskID = g_pages[window.g_curIndex]["序号"];
      if (rsq.length==0)
      {
        alert("不能为空");
        return;
      }
      if (window.XMLHttpRequest)
      {// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp=new XMLHttpRequest();
      }
      else
      {// code for IE6, IE5
        xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
      }
        xmlhttp.onreadystatechange=function()
      {
      if (xmlhttp.readyState==4 && xmlhttp.status==200)
        {
          if (xmlhttp.responseText == "True")
          {
            g_pages[window.g_curIndex]["IsMatch"] = "True";
            document.getElementById("CheckBtn").disabled = true;
            document.getElementById("RealStockQty").disabled = true;
            onClickPageBtn(1);
          }
          else
          {
            alert("数据不一致，请重新检查");
          }
          g_pages[window.g_curIndex]["CheckedTimes"] += 1;
          if (g_pages[window.g_curIndex]["CheckedTimes"] > 2)
          {
            alert("是否确认最终数量？");
            document.getElementById("CheckBtn").disabled = true;
            document.getElementById("RealStockQty").disabled = true;
            onClickPageBtn(1);
          }
          
        }
      }
    xmlhttp.open("GET","ajax_check_storage_qty.asp?rsq="+rsq+"&taskID="+taskID,true);
    xmlhttp.send();
  }
  
  function recoverTaskStatus()
  {
    var localStatus=getCookie(window.g_staffNo);
    var taskID;
    var i;
    if (typeof(localStatus) === 'undefined' || localStatus == null)
    {
      return false;
    }
    var lsObj = eval('('+localStatus+')');
    for (i=0; i<window.g_totalCnt;++i)
    {
      taskID = window.g_pages[i]["序号"];
      if (typeof(lsObj[taskID]) === 'undefined' || lsObj[taskID]["PartsCode"] != window.g_pages[i]["PartsCode"])
      {
        return;
      }
    }
    for (i=0; i<window.g_totalCnt;++i)
    {
      taskID = window.g_pages[i]["序号"];
      window.g_pages[i]["RealStockQty"]=lsObj[taskID]["RealStockQty"];
    }
  }
    /**
  * json对象转字符串形式
  */
  function json2str(o)
  {
    var arr = [];
    var fmt = function(s) {
      if (typeof s == 'object' && s != null) return json2str(s);
      return /^(string|number)$/.test(typeof s) ? "'" + s + "'" : s;
    }
    for (var i in o) arr.push("'" + i + "':" + fmt(o[i]));
    return '{' + arr.join(',') + '}';
  } 
  
	function getCookie(name)
	{
	    var cookieStr = window.document.cookie;
	    var reg = new RegExp(name+"=[a-zA-Z0-9%]*;", "m");
	    var cookieArr = cookieStr.match(reg);
	    if (cookieArr == null)
	    {
	      return null;
	    }
	    var cookieVal = cookieArr[0].split("=");
	    if(cookieVal[0] == name)
	    {
	        return unescape(cookieVal[1].slice(0, -1));
	    }
	    
	    return null;
	}
	
	//创建cookie
  function setCookie(name, value, expireday)
  {
  	var exp = new Date();
  	exp.setTime(exp.getTime() + expireday*60*60*1000); //设置cookie的期限
  	window.document.cookie = name+"="+escape(value)+"; expires"+"="+exp.toGMTString();//创建cookie
  }
  
  function deleteCookie(name)
  { 
    var exp = new Date(); 
    exp.setTime(exp.getTime() - 1); 
    var cval = getCookie(name);
    if (cval == null)
    {
      return;
    }
    window.document.cookie = name + "=" + cval + "; expires=" + exp.toGMTString();
  }
</script>
</body>
</html>
