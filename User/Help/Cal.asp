<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/UserTopic.asp -->
<%
DEF_BBS_HomeUrl = "../../"
SiteHead(DEF_SiteNameString & " - �û���Ա�� - ��̳����")

UserTopicTopInfo
DisplayUserNavigate("��̳����")%>
<br><br>
<%Help_Cal
UserTopicBottomInfo
sitebottom
If GBL_ShowBottomSure = 1 Then Response.Write GBL_SiteBottomString

Sub Help_Cal
%>
<SCRIPT language=JScript>
<!--
var Today = new Date(<%=year(DEF_Now)%>,<%=month(DEF_Now)-1%>,<%=day(DEF_Now)%>,<%=hour(DEF_Now)%>,<%=minute(DEF_Now)%>,<%=second(DEF_Now)%>);
//-->
</SCRIPT>
<script languange="Javascript" src="Cal.js"></script>
<STYLE>.todyaColor {
	BACKGROUND-COLOR: aqua
}
</STYLE>

<BODY onload=initialize() onunload=terminate()>
<SCRIPT language=JavaScript><!--
   if(navigator.appName == "Netscape" || parseInt(navigator.appVersion) < 4)
   document.write("<h1>���������޷�ִ�д˳���</h1>�˳������� IE4 �Ժ�İ汾����ִ��!!")
//--></SCRIPT>

<DIV id=detail style="Z-INDEX: 3; FILTER: shadow(color=#333333,direction=135); WIDTH: 140px; POSITION: absolute; HEIGHT: 120px"></DIV>
<CENTER>
<TABLE border=0 CELLPADDING=0 CELLSPACING=0 width=100%>
  <TBODY>
  <TR><!------------------------------ ����ʱ�� ----------------------------------->
    <FORM name=WorldClock>
    <TD vAlign=top align=middle width=220>����ʱ��<BR><SPAN id=LocalTime 
      style="FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: Arial">0000��0��0��(��)�� 
      00:00:00</SPAN> 
      <P><SPAN id=City style="WIDTH: 150px;">�й�</SPAN> 
      <BR><SPAN id=GlobeTime>0000��0��0��(��)�� 00:00:00</SPAN><BR>
      <TABLE sCELLPADDING=0 CELLSPACING=0 width=100%>
        <TBODY>
        <TR>
          <TD align=middle>
            <DIV id=map style="FILTER: Light; OVERFLOW: hidden; WIDTH: 190px; HEIGHT: 120px; BACKGROUND-COLOR: mediumblue"><FONT id=world style="FONT-SIZE: 185px; LEFT: 0px; COLOR: green; FONT-FAMILY: Webdings; POSITION: relative; TOP: -26px">��</FONT> 
            </DIV></TD></TR></TBODY></TABLE><BR><SELECT Style="WIDTH: 190px; BACKGROUND-COLOR: <%=DEF_BBS_TableHeadColor%>" onchange=chContinent() name=continentMenu></SELECT><BR><SELECT Style="WIDTH: 190px; BACKGROUND-COLOR: <%=DEF_BBS_TableHeadColor%>" onchange=chCountry() name=countryMenu></SELECT></P></TD></FORM>
		<FORM name=CLD>
    <TD align=middle valign=top><img src=../../images/null.gif width=5 height=4>
    <TD align=middle valign=top width=80%>
      <TABLE border=0 CELLPADDING=1 CELLSPACING=1 width=100% bgcolor=<%=DEF_BBS_DarkColor%> class=TBone>
        <TBODY>
        <TR>
          <TD bgColor=<%=DEF_BBS_LightDarkColor%> colSpan=7 height=25><FONT color=#ffffff size=2>��Ԫ<SELECT onchange=changeCld() name=SY style="BACKGROUND-COLOR: <%=DEF_BBS_LightestColor%>" class=TBBG9> 
              <SCRIPT language=JavaScript><!--
          for(i=1900;i<2101;i++) document.write('<option>'+i)
            //--></SCRIPT>
            </SELECT>��<SELECT onchange=changeCld() 
            name=SM style="BACKGROUND-COLOR: <%=DEF_BBS_LightestColor%>" class=TBBG9> 
              <SCRIPT language=JavaScript><!--
            for(i=1;i<13;i++) document.write('<option>'+i)
            //--></SCRIPT>
            </SELECT>��</FONT> <FONT id=GZ face=�꿬�� color=#ffffff></FONT><BR></TD></TR>
        <TR align=middle bgColor=<%=DEF_BBS_Color%> class=TBthree>
          <TD width=54>��</TD>
          <TD width=54>һ</TD>
          <TD width=54>��</TD>
          <TD width=50>��</TD>
          <TD width=54>��</TD>
          <TD width=54>��</TD>
          <TD width=54>��</TD></TR>
        <SCRIPT language=JavaScript><!--
            var gNum, color1, color2;

            // ��������ɫ
            switch (conWeekend) {
            case 1:
               color1 = 'black';
               color2 = color1;
               break;
            case 2:
               color1 = 'green';
               color2 = color1;
               break;
            case 3:
               color1 = 'red';
               color2 = color1;
               break;
            default :
               color1 = 'green';
               color2 = 'red';
            }

            for(i=0;i<6;i++) {
               document.write('<tr align=center bgcolor=<%=DEF_BBS_LightestColor%> class=TBBG9>')
               for(j=0;j<7;j++) {
                  gNum = i*7+j
                  document.write('<td id="GD' + gNum +'" onMouseOver="mOvr(' + gNum +')" onMouseOut="mOut()"><font id="SD' + gNum +'" face="Arial Black"')
                  if(j == 0) document.write(' color=red class=RedFont')
                  if(j == 6) {
                     if(i%2==1) document.write(' color='+color2)
                        else document.write(' color='+color1)
                  }
                  document.write(' TITLE=""> </font><br><font id="LD' + gNum + '" style="font-size:9pt"> </font></td>')
               }
               document.write('</tr>')
            }
            //--></SCRIPT>
        </TBODY></TABLE></TD>
        <td width=20></td>
    <tr><td></td><td></td><td>
    ��<BUTTON onclick="pushBtm('YD')" class=fmbtn><B>��</B></BUTTON>
    <BUTTON onclick="pushBtm('YU')" class=fmbtn><B>��</B></BUTTON> 
      ��
      <BUTTON onclick="pushBtm('MD')" class=fmbtn><B>��</B></BUTTON>
      <BUTTON onclick="pushBtm('MU')" class=fmbtn><B>��</B></BUTTON> 
      &nbsp; <BUTTON onclick="pushBtm('')" class=fmbtn>����</BUTTON> 
      </TD><td width=20></td><TD></TR></FORM></TBODY></TABLE><FONT color=#ffffff>
</FONT><BR></CENTER></FONT></FONT></BODY>
<%
end Sub%>
