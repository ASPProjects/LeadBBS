<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/UserTopic.asp -->
<%
DEF_BBS_HomeUrl = "../../"

Main

Sub Main

	Select Case Left(Request.QueryString,4)
		Case "icon"
			BBS_SiteHead DEF_SiteNameString & " - ��̳����",0,"<span class=navigate_string_step>��̳����</span>"
			UserTopicTopInfo("help")
			Help_UbbIcon
		Case "colo"
			BBS_SiteHead DEF_SiteNameString & " - ��ɫ��",0,"<span class=navigate_string_step>��ɫ��</span>"
			UserTopicTopInfo("help")
			Help_UbbColor
		Case else
			BBS_SiteHead DEF_SiteNameString & " - UBB��������",0,"<span class=navigate_string_step>UBB��������</span>"
			UserTopicTopInfo("help")
			Help_UbbCode
	End Select
	UserTopicBottomInfo
	sitebottom

End Sub

Function Help_UbbColor%>

<img src=<%=DEF_BBS_HomeUrl%>images/others/colortable.GIF>

<%End Function

Sub Help_UbbCode

	%>

<div class=title>ʲô��UBB���룿 </div>
<p>UBB������HTML��һ�����֣���Ultimate Bulletin Board(�����һ��BBS����)���õ�һ�������TAG����Ҳ���Ѿ���������Ϥ�ˡ�UBB����ܼ򵥣����ܺ��٣�����������Tag�﷨���ʵ�ַǳ����ף��������ǵ���վ���������ִ��룬�Է�������ʹ����ʾͼƬ/����/�Ӵ�����ȳ������ܡ� 

<div class=title>UBB�������ʵ����ЩHTML�Ĺ��ܣ�������ʹ�����Ӻͼ��ɣ�</div>
<OL class=helpOL>
<li>�����ַ������Լ��볬�����ӣ��������Ӿ����ַ������������
<div class=value2><span class=redfont>[URL]</span><a href=http://www.%4c%65%61%64%42%42%53.%63%6f%6d/>http://www.asph.net/</a><span class=redfont>[/URL] 
  </span>
</div>
<div class=value2>[URL=</span><a href=http://%4c%65%61%64%42%42%53.%63%6f%6d/>http://www.asph.net/</a><span class=redfont>]</span>LeadBBS<span class=redfont>[/URL]</span>
</div>
<li>��ʾΪ����Ч��
<div class=value2><span class=redfont>[B]</span>����<span class=redfont>[/B]</span></div>
<li>��ʾΪб��Ч��
<div class=value2><span class=redfont>[I]</span>����<span class=redfont>[/I]</span></div>
<li>��ʾΪ�»���Ч��
<div class=value2><span class=redfont>[U]</span>����<span class=redfont>[/U]</span></div>
<li>����λ�ÿ���
<p>�����ֵ�λ�ÿ��������������Ҫ���ַ���<span class=bluefont>center</span>λ��<span class=redfont>center</span>��ʾ���У�<span class=redfont>left</span>��ʾ����<span class=redfont>right</span>��ʾ���ң�<span class=redfont>justify</span>��ʾ���˶���
<div class=value2><span class=redfont>[ALIGN=<span class=bluefont>center</span>]</span><br>
  ���ֶ���<br>
  <span class=redfont>[/ALIGN]</span></div>
<li>�ּ��(�и�)����
<p>�и�ֵ������<span class=bluefont>normal</span>�����Ǿ�����и���ֵ������<span class=bluefont>150%</span>��<span class=bluefont>1.5</span>��<span class=bluefont>24pt</span>
<div class=value2><span class=redfont>[LINE-HEIGHT=<span class=bluefont>150%</span>]</span>
  ���ֶ���
  <span class=redfont>[/LINE-HEIGHT]</span></div>
<li>�����ʼ����������ַ������ԣ��������Ӿ����ַ������������<br>

<div class=value2><span class=redfont>[EMAIL]</span>webmaster@LeadBBS.com<span class=redfont>[/EMAIL] 
  </span></div>
<div class=value2><span class=redfont>[EMAIL=</span>webmaster@LeadBBS.com<span class=redfont>]</span>LeadBBS<span class=redfont>[/EMAIL]</span></div>
<li>����ͼƬ
<div class=value2><span class=redfont>[IMG]</span>http://www.LeadBBS.com/images/flag.gif<span class=redfont>[/IMG]</span><br>

<br>����ͼƬ��ָ�����뷽ʽ���߿��С�����뷽ʽ�� <span class=bluefont>absmiddle left right top middle<br> bottom absbottom baseline texttop</span><p>
<span class=redfont>[IMG=<span class=bluefont>2</span>,<span class=bluefont>center</span>]</span>http://www.LeadBBS.com/images/flag.gif<span class=redfont>[/IMG]</span><br>
<span class=redfont>[IMG=<span class=bluefont>2,���뷽ʽ,�߶�,���</span>]</span>http://www.LeadBBS.com/images/flag.gif<span class=redfont>[/IMG]</span><br>

<li>����MicroMedia��Flash
<div class=value2><span class=redfont>[Flash]</span>http://www.test.com/flag.swf<span class=redfont>[/Flash]</span></div>
<div class=value2><span class=redfont>[Flash=<span class=bluefont>���,�߶�</span>]</span>http://www.test.com/flag.swf<span class=redfont>[/Flash]</span></div>
<li>ʵ�ִ��빦�ܣ���ʹUBB���뷽ʽ��Ҳ������ʾUBB����
<div class=value2><span class=redfont>[CODE]<br>
  </span>���ֶ���<br>
  <span class=redfont>[/CODE]</span></div>
<li>����Ч�����ñ�����
<div class=value2><span class=redfont>[QUOTE]</span><br>
  ���ö���<br>
  <span class=redfont>[/QUOTE]</span></div>
<li>ʵ��HTMLĿ¼Ч��
<div class=value2><span class=redfont>[UL]</span>����<span class=redfont>[/UL]</span> - 
  �൱��html�е�&lt;UL&gt;���ܣ������Ű�<br>
  <span class=redfont>[OL]</span>����<span class=redfont>[/OL]</span> 
  - �൱��html�е�&lt;OL&gt;�����������ֱ�ŵ�Ч��<br>
  <span class=redfont>[LI]</span>����<span class=redfont>[/LI]</span> 
  - �൱��html�е�&lt;li&gt;�������ϱ�ǩ����ʹ��
<li>ʵ�����ַ���Ч��(�����)���൱��html�е�&lt;marquee&gt;
<div class=value2><span class=redfont>[FLY]</span>����<span class=redfont>[/FLY]</span></div>
<li>���뵥Ԫ��
<div class=value2><span class=redfont>[HR]</span>
<li>ʵ�����ַ�����Ч��GLOW����������Ϊ���롢��ɫ�ͱ߽��С
<div class=value2><span class=redfont>[GLOW=</span><span class=bluefont>1,RED,2</span><span class=redfont>]</span>����<span class=redfont>[/GLOW]</span></div>
<li>ʵ��������Ӱ��Ч��SHADOW����������Ϊ���롢��ɫ�ͱ߽��С
<div class=value2><span class=redfont>[SHADOW=</span><span class=bluefont>1,RED,2</span><span class=redfont>]</span>����<span class=redfont>[/SHADOW]</span></div>
<li>ʵ��������ɫ�ı�
<div class=value2><span class=redfont>[COLOR=</span><span class=bluefont>��ɫ</span><span class=redfont>]</span>����<span class=redfont>[/COLOR]</span></div>
<li>ʵ�����ִ�С�ı�
<div class=value2><span class=redfont>[SIZE=</span><span class=bluefont>����1-9��CSS����ĳߴ綨�崮</span><span class=redfont>]</span>����<span class=redfont>[/SIZE]</span></div>
<li>ʵ����������ת��
<div class=value2><span class=redfont>[FACE=</span><span class=bluefont>����</span><span class=redfont>]</span>����<span class=redfont>[/FACE]</span></div>

<li>�����л���
<div class=value2><span class=redfont>[STRIKE]</span>����<span class=redfont>[/STRIKE]</span></div>

<li>����RealPlayer��ʽ��rm�ļ����м������Ϊ��Ⱥͳ���
<div class=value2><span class=redfont>[RM=<span class=bluefont>���</span>,<span class=bluefont>�߶�</span>]</span>http://....<span class=redfont>[/RM]</span></div>

<li>����ΪMidia Player��ʽ���ļ����м������Ϊ��Ⱥͳ���
<div class=value2><span class=redfont>[MP=<span class=bluefont>���</span>,<span class=bluefont>�߶�</span>]</span>http://....<span class=redfont>[/MP]</span></div>

<li>�ϱ�����
<div class=value2><span class=redfont>[sup]</span>����<span class=redfont>[/sup]</span>��Ч����LeadBBS<sup>2</sup>

<li>�±�����
<div class=value2><span class=redfont>[sub]</span>����<span class=redfont>[/sub]</span>��Ч����LeadBBS<sub>2</sub>

<li>ָ��������ɫ��������ɫ
<div class=value2><span class=redfont>[BGCOLOR=<span class=bluefont>������ɫ</span>]</span>����<span class=redfont>[/BGCOLOR]</span><br>
<span class=redfont>[BGCOLOR=<span class=bluefont>������ɫ</span>,<span class=bluefont>������ɫ</span>]</span>����<span class=redfont>[/BGCOLOR]</span></div>

<li>���뱳������
<div class=value2><span class=redfont>[SOUND]</span>���������ļ���ַ<span class=redfont>[/SOUND]</span></div>

<li>������Ŀ��
<div class=value2><span class=redfont>[FIELDSET=<span class=bluefont>����</span>]</span>����<span class=redfont>[/FIELDSET]</span></div>

<li>������˸Ч��
<div class=value2><span class=redfont>[LIGHT]</span>��˸����<span class=redfont>[/LIGHT]</span></div>

<li>������
<div class=value2><span class=redfont>[TABLE][TR][TD]</span>����<span class=redfont>[/TD][/TR][/TABLE]</span></div>

<div class=value2>���븴�ӵı��8����������ָ����ȫ</div>
<div class=value2><span class=redfont>[TABLE=<span class=bluefont>�߿�ɫ</span>,<span class=bluefont>��Ԫ���</span>,<span class=bluefont>��Ԫ�߾�</span>,<span class=bluefont>����</span>,<span class=bluefont>���뷽ʽ</span>,<span class=bluefont>����ɫ</span>,<span class=bluefont>�߿��ϸ</span>,<span class=bluefont>����ͼƬ</span>][TR][TD]</span>����<span class=redfont>[/TD][/TR][/TABLE]</span></div>
<div class=value2>TD��ǩ�������Ķ��壺[TD=��,��,����ɫ]������ɫ��ָ����Ҳ�ɲ�ָ��</div>

<li>�ѱ��Ÿ�ʽ����ͬ��HTML�е�&lt;PRE&gt;��ǩ
<div class=value2><span class=redfont>[PRE]</span>����<span class=redfont>[/PRE]</span></div>

<li>�����LRC��ʵĸ���
<div class=value2><span class=redfont>[LRC=<span class=bluefont>�����ļ���ַ</span>]</span><span class=bluefont>LRC�ļ���ַ��LRC�ļ�����</span><span class=redfont>[/LRC]</span></div>

<li>�����۵�������
<div class=value2><span class=redfont>[collapse<span class=bluefont>=�۵���ʾ��Ϣ(�����ѡ)</span>]</span><span class=bluefont>������۵�����</span><span class=redfont>[/collapse]</span></div>

<%End Sub

Sub Help_UbbIcon

	Dim N
	%><div class=title>��̳����</div>
	<br>
	<table border=0 cellpadding=0 cellspacing=0 class=blanktable>
	<tr>
		<td>���</td>
		<td>����</td>
		<td>����</td>
	</tr><%
	For n = 0 to DEF_UBBiconNumber - 1
		%>
	<tr>
		<td><%=Right(" " & n + 1,2)%></td>
		<td><img src="<%=DEF_BBS_HomeUrl%>images/UBBicon/em<%=Right("0" & n + 1,2)%>.GIF" width=20 height=20 align=middle border=0></td>
		<td>[em<%=Right("0" & n + 1,2)%>]</td>
	</tr>
		<%
	Next
	%>
	</table>
	<%

End Sub%>