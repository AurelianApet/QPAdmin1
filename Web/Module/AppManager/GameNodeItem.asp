﻿<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../Cls_Page.asp"-->
<!--#include file="../../conn.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="../../Css/layout.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../../Js/common.js"></script>    
    <script type="text/javascript" src="../../Js/Check.js"></script>  
    <script type="text/javascript" src="../../Js/comm.js"></script>   
    <script type="text/javascript" src="../../Js/Sort.js"></script> 
</head>
<body>
    <!-- 头部菜单 Start -->
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">你当前位置：系统维护 - 游戏管理</td>
        </tr>
    </table>
    <!-- 头部菜单 End -->
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="Tmg7">
        <tr>
            <td height="28">
                <ul>
                    <li class="tab2" onclick="openWindowSelf('GameGameItem.asp')">模块</li>
	                <li class="tab2" onclick="openWindowSelf('GameTypeItem.asp')">类型</li>
	                <li class="tab2" onclick="openWindowSelf('GameKindItem.asp')">游戏</li>
	                <li class="tab1">节点</li>
	                <li class="tab2" onclick="openWindowSelf('GamePageItem.asp')">自定义页</li>
                </ul>
            </td>
        </tr>
    </table> 
    <% 
        Call ConnectGame(QPPlatformDB)
        Select case lcase(Request("action"))
            case "nodeinfo"
            Call GetNodeInfo()
            case "save"
            Call SaveNodeInfo()
            case else
            Call Main()
        End Select
        Call CloseGame()
        
        '保存
        Sub SaveNodeInfo()
            '判断重复
            Dim ret
            Set ret = GameConn.execute("select * from GameNodeItem where NodeID<>'"&Request("id")&"' AND NodeID='"&CxGame.GetInfo(1,"in_NodeID")&"'")
            If Not ret.eof Then
                Call CxGame.ShowInfo2("节点标识重复，操作失败！","GameNodeItem.asp",1200)
                Exit Sub
            End If
            
            '保存开始
            Dim rs,sql
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GameNodeItem where NodeID='"&Request("id")&"'"
            rs.Open sql,GameConn,1,3
            If rs.eof Then
                rs.addnew
            End If
            rs("NodeID") = CxGame.GetInfo(1,"in_NodeID")
            rs("KindID") = CxGame.GetInfo(1,"in_KindID")
            rs("JoinID") = CxGame.GetInfo(1,"in_JoinID")
            rs("SortID") = CxGame.GetInfo(1,"in_SortID")
            rs("NodeName") = CxGame.GetInfo(0,"in_NodeName")
            rs("Nullity") = CxGame.GetInfo(0,"in_Nullity")
            rs.update
            If Request("id")<>"" Then
                Call CxGame.ShowInfo2("修改成功！","GameNodeItem.asp?page="&Request("page")&"",1200)
            Else
                Call CxGame.ShowInfo2("新增成功！","GameNodeItem.asp",1200)
            End If
            rs.close()
            Set rs = Nothing   
        End Sub
        
        '删除
        Sub Delete(lID)
            Dim ID
            ID = lID
            GameConn.execute "delete GameNodeItem where NodeID='"&ID&"'"            
        End Sub
        
        Sub Main()   
            Dim cidArray, lLoop
            cidArray = Split(Request("cid"),",")
            For lLoop=0 To UBound(cidArray)
                Call Delete(cidArray(lLoop))
            Next
            
            Dim rs,nav,Page,i
            Dim lCount, queryCondition, OrderStr
             If Request.QueryString("orders")<>""And Request.QueryString("OrderType")<>"" Then
                If Request.QueryString("OrderType")<>"-102" Then
                    OrderStr=Replace(Request.QueryString("orders")," ","+")&" Asc "
                End If
                If Request.QueryString("OrderType")<>"-101" Then
                    OrderStr=Replace(Request.QueryString("orders")," ","+")&" Desc "
                End If
            Else 
                OrderStr = "SortID Asc"
            End If
            
            '==============================================================================================================
            '执行存储过程
                    
            Set Page = new Cls_Page      '创建对象
            Set Page.Conn = GameConn     '得到数据库连接对象        
            With Page
                .PageSize = GetPageSize					'每页记录条数
                .PageParm = "page"				'页参数
                '.PageIndex = 10				'当前页，可选参数，一般是生成静态时需要
	            .Database = "mssql"				'数据库类型,AC为access,MSSQL为sqlserver2000存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	            .Pkey="NodeID"					'主键
	            .Field="NodeID,KindID,JoinID,SortID,NodeName,Nullity"	'字段
	            .Table="GameNodeItem"			'表名
	            .Condition=queryCondition		'条件,不需要where
	            .OrderBy=OrderStr				'排序,不需要order by,需要asc或者desc
	            .RecordCount = 0				'总记录数，可以外部赋值，0不保存（适合搜索），-1存为session，-2存为cookies，-3存为applacation

	            .NumericJump = 5                '数字上下页个数，可选参数，默认为3，负数为跳转个数，0为显示所有
	            .Template = "总记录数：{$RecordCount} 总页数：{$PageCount} 每页记录数：{$PageSize} 当前页数：{$PageIndex} {$FirstPage} {$PreviousPage} {$NumericPage} {$NextPage} {$LastPage} {$InputPage} {$SelectPage}" '整体模板，可选参数，有默认值
	            .FirstPage = "首页"             '可选参数，有默认值
	            .PreviousPage = "上一页"        '可选参数，有默认值
	            .NextPage = "下一页"            '可选参数，有默认值
	            .LastPage = "尾页"              '可选参数，有默认值
	            .NumericPage = " {$PageNum} "   '数字分页部分模板，可选参数，有默认值
            End With
            
            rs = Page.ResultSet()       '记录集
            lCount = Page.RowCount()    '可选，输出总记录数
            nav = Page.Nav()            '分页样式
            
            Page = Null
            Set Page = Nothing
            '==============================================================================================================
    %>
    <script type="text/javascript">
        function CheckFormList()
        {
            if(!confirm("确定要执行选定的操作吗？"))
            {
                return false;
            }
            else
            {
                var cid="";
                var cids = GelTags("input");                
                for(var i=0;i<cids.length;i++)
                {
                    if(cids[i].checked)
                    {
                        if(cids[i].name == "cid")
                            cid+=cids[i].value+",";
                    }
                }       
                if(cid=="")
                {
                    showError("未选中任何数据");
                    return false;
                }
            }
        }
    </script>
    <form name="myformList" method="post" action='' onsubmit="return CheckFormList();">
        <% If Request("Action") = "DeleteAll" Then %>
        <script type="text/javascript">
            showInfo("删除成功")
        </script>
        <% End If %>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td height="39" class="titleOpBg">
	                <input type="button" value="新建" class="btn wd1" onclick="Redirect('?action=nodeinfo')" />
                    <input class="btnLine" type="button" />
                    <input type="submit" value="删除" class="btn wd1" />            
                    <input type="hidden" name="Action" value="DeleteAll" />         
                </td>
            </tr>
        </table> 
        <div id="content">
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
                <tr align="center" class="bold">
                    <th class="listTitle1" width='38' align='center'><input type="checkbox" name="chkAll" onclick="SelectAll(this.checked);" /></th>
                    <th class="listTitle2" width="80"><a href="" class="l1"  onclick="GetOrder('NodeID',this)">节点标识</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('SortID',this)">节点名称</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('KindID',this)">游戏名称</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('JoinID',this)">挂接</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('NodeName',this)">排序</a></th>
                    <th class="listTitle"><a class="l1"  href="" onclick="GetOrder('Nullity',this)">禁用状态</a></th>
                </tr>
                <% 
                    Dim NodeID,className
                    If IsNull(rs) Then
                        Response.Write("<tr class='tdbg'><td colspan='100' align='center'><br>没有任何信息!<br><br></td></tr>")
                    Else
                    
                    For i=0 To Ubound(rs,2)
                    NodeID = rs(0,i)   
                    If i Mod 2 = 0 Then
                        className="list"        
                    Else
                        className="listBg"      
                    End If              
                %>
                <tr class="<%=className %>" onmouseover="currentcolor=this.style.backgroundColor;this.style.backgroundColor='#caebfc';this.style.cursor='pointer';"
                    onmouseout="this.style.backgroundColor=currentcolor">
                    <td><input name='cid' type='checkbox' value='<%=NodeID%>'/></td>                  
                    <td onclick="Redirect('?action=nodeinfo&id=<%=NodeID %>&page=<%=Request("page") %>')"><%=rs(0,i) %></td>    
                    <td onclick="Redirect('?action=nodeinfo&id=<%=NodeID %>&page=<%=Request("page") %>')"><%=rs(4,i) %></td>    
                    <td onclick="Redirect('?action=nodeinfo&id=<%=NodeID %>&page=<%=Request("page") %>')"><%=CxGame.GetKindNameByKindID(rs(1,i)) %></td>    
                    <td onclick="Redirect('?action=nodeinfo&id=<%=NodeID %>&page=<%=Request("page") %>')"><%=CxGame.GetNodeNameByNodeID(rs(2,i)) %></td>      
                    <td onclick="Redirect('?action=nodeinfo&id=<%=NodeID %>&page=<%=Request("page") %>')"><%=rs(3,i) %></td>   
                    <td onclick="Redirect('?action=nodeinfo&id=<%=NodeID %>&page=<%=Request("page") %>')"><%=CxGame.GetNullityInfo(rs(5,i)) %></td>            
                </tr>
                <% 
                    Next
                    End If
                %>
            </table>
        </div>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="listTitleBg"><span>选择：</span>&nbsp;<a class="l1" href="javascript:SelectAll(true);">全部</a>&nbsp;-&nbsp;<a class="l1" href="javascript:SelectAll(false);">无</a></td>
                <td class="page" align="right"><%Response.Write nav%></td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="OpList">
            <tr>
                <td height="39" class="titleOpBg">
	                <input type="button" value="新建" class="btn wd1" onclick="Redirect('?action=nodeinfo')" />
                    <input class="btnLine" type="button" />
                    <input type="submit" value="删除" class="btn wd1" />            
                </td>
            </tr>
        </table>
    </form>
    <% 
        End Sub
        
        Sub GetNodeInfo()
            Dim rs,sql,titleInfo
            Dim nodeID,kindID,joinID,sortID,nodeName,nullity
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GameNodeItem where NodeID='"&Request("id")&"'"
            rs.Open sql,GameConn,1,3
            If rs.Bof And rs.Eof Then
                titleInfo = "新增节点"
                nodeID=0
            Else
                titleInfo = "修改节点"
                nodeID = rs("NodeID")
                kindID = rs("KindID")
                joinID = rs("JoinID")
                sortID = rs("SortID")
                nodeName = rs("NodeName")                
                nullity = rs("Nullity")
            End If       
    %>
    <script type="text/javascript">
        function CheckFormInfo()
        {
            var nodeID = document.myFormInfo.in_NodeID.value;
            var nodeName = document.myFormInfo.in_NodeName.value;
            if(nodeID.trim()=="")
            {
                alert("节点标识不能为空！");
                return false;
            }
            if(IsPositiveInt(nodeID.trim())==false)
            {
                alert("节点标识非数字字符或者超过允许大小！");
                return false;
            }
            if(nodeName.trim()=="")
            {
                alert("节点名称不能为空！");
                return false;
            }
            else if(len(nodeName.trim())>32)
            {
                alert("节点名称字符长度不可超过32个字符！");
                return false;
            }
        }
    </script>
    <form name="myFormInfo" method="post" action='?action=save&id=<%=Request("id") %>&page=<%=Request("page") %>' onsubmit="return CheckFormInfo()">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GameNodeItem.asp')" />                                       
                    <input class="btnLine" type="button" />  
                    <input type="submit" value="保存" class="btn wd1" />  
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="listBg2">
            <tr>
                <td height="35" colspan="2" class="f14 bold Lpd10 Rpd10"><div class="hg3  pd7"><%=titleInfo %></div></td>
            </tr>
            <tr>
                <td class="listTdLeft">节点标识：</td>
                <td>
                    <% If nodeID=0 Then%>
                    <input name="in_NodeID" type="text" class="text" value="<%=CxGame.GetMaxNodeID()+1 %>" />  
                    <% Else%>      
                    <input name="in_NodeID" type="text" class="text" readonly="readonly" value="<%=nodeID %>" />  
                    <% End If%>             
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">节点名称：</td>
                <td>
                    <input name="in_NodeName" type="text" class="text" value="<%=nodeName %>" />                
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">游戏名称：</td>
                <td>
                    <select name="in_KindID" style="width:157px;">
                    <% 
                        Dim ArrayKind,i
                        ArrayKind = CxGame.GetKindList()
                        For i=0 To Ubound(ArrayKind)    
                        If kindID=ArrayKind(i,0) Then                                                      
                    %>
                    <option value="<%=ArrayKind(i,0) %>" selected="selected"><%=ArrayKind(i,1) %></option>
                    <% Else%>
                    <option value="<%=ArrayKind(i,0) %>"><%=ArrayKind(i,1) %></option>
                    <% 
                        End If                        
                        Next  
                        Set ArrayKind = nothing
                    %>
                    </select>                   
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">挂接：</td>
                <td>
                    <select name="in_JoinID" style="width:157px;">
                    <option value="0">无挂接</option>
                    <% 
                        Dim ArrayNode,j
                        ArrayNode = CxGame.GetNodeList()
                        For j=0 To Ubound(ArrayNode)    
                        If joinID=ArrayNode(j,0) Then    
                    %>
                    <option value="<%=ArrayNode(j,0) %>" selected="selected"><%=ArrayNode(j,1) %></option>
                    <% Else%>
                    <option value="<%=ArrayNode(j,0) %>"><%=ArrayNode(j,1) %></option>
                    <% 
                        End If
                        Next 
                        Set ArrayNode = nothing
                    %>
                    </select>             
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">排序：</td>
                <td>
                    <input name="in_SortID" type="text" class="text" value="<%=sortID %>" />                  
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">禁用状态：</td>
                <td>
                    <input name="in_Nullity" id="in_Nullity0" type="radio" value="0"<% If nullity=0 Then %> checked="checked"<% End If %> /><label for="in_Nullity0">正常</label>
                    <input name="in_Nullity" id="in_Nullity1" type="radio" value="1"<% If nullity=1 Then %> checked="checked"<% End If %> /><label for="in_Nullity1">禁用</label>              
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GameNodeItem.asp')" />                                       
                    <input class="btnLine" type="button" />  
                    <input type="submit" value="保存" class="btn wd1" />  
                </td>
            </tr>
        </table>
    </form>
    <%
        rs.Close()
        Set rs=nothing
        End Sub
    %>
</body>
</html>
