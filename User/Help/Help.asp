<!-- #include file=../../inc/BBSsetup.asp -->
<!-- #include file=../../inc/User_Setup.asp -->
<!-- #include file=../../inc/Board_popfun.asp -->
<!-- #include file=../inc/UserTopic.asp -->
<%
DEF_BBS_HomeUrl = "../../"

Main

Sub Main

	Select Case Left(Request.QueryString,3)
		Case else
			BBS_SiteHead DEF_SiteNameString & " - �����ĵ�����",0,"<span class=navigate_string_step>�����ĵ�����</span>"
			UserTopicTopInfo("help")
			Help_Index
	End Select
	UserTopicBottomInfo
	sitebottom

End Sub

Sub Help_Index%>

<div class=title>�����ĵ�����</div>
<ol>
<li><a href=#001>��վ�û�����������</a></li>
<li><a href=#FullSearchHelp>ȫ�ļ���ʹ�ð���</a></li>
<li><a href=#003>�������</a></li>
<li><a href=#004>�����������</a></li>
<li><a href=#005>��̳�����û����<u><%=DEF_PointsName(9)%></u>һ����</a></li>
<li><a href=#006>��̳����ͼ��</a></li>
<li><a href=#007>��̳����ͼ��</a></li>
<li><a href=#008>��̳���ܼ�������ͼ��</a></li>
</ol>
<hr class=splitline>

<div class=title>��վ�û�����������</div>
<OL class=helpOL>
<li>
	Ϊ��Ҫע���Ϊ��վ�û���
	<div class=value2>ע���Ϊ��վ�û���������ȫ���ܱ�վ�ṩ��ȫ�����񣬲��ܹ������뱾վ����<br>
	���û����и��õĽ�����</div>
</li>
<li>Ϊʲô�ҵ�¼���κ����ʾ����¼̫Ƶ��������Ϣ���ٵ�¼����ϵ����Ա! ������ʾ��<br>
	<div class=value2>
	Ϊ���û��İ�ȫ�����û����������¼��������������,�û�������5���Ӻ�Ϳ������µ�¼����������������ڷǷ��������¼��Ŀǰʹ���ŵ��û��ܹ���������ʹ���ʺš�
	</div>
</li>
<li>�û�<%=DEF_PointsName(0)%>����μ���ģ�
<div class=value2><%=DEF_PointsName(0)%>��Դ�����¼������棺</div>
	<div class=value2>��̳����һƪ�������ӣ���<%=DEF_BBS_AnnouncePoints*2%><%=DEF_PointsName(0)%><br>
  	��̳����һƪ�ظ����ӣ���<%=DEF_BBS_AnnouncePoints%><%=DEF_PointsName(0)%><br>
	��̳һƪ���ӳ�Ϊ��������������<%=DEF_BBS_MakeGoodAnnouncePoints%><%=DEF_PointsName(0)%>
	</div>
<li>
	<a name=#UserLevel>�û�<%=DEF_PointsName(3)%>�б�</a>
	<div class=value2>
<table border="0" cellspacing="0" cellpadding="0" class=table_in>
  <tr>
    <td>
        <tr class=tbinhead> 
          <td><div class=value><%=DEF_PointsName(3)%></div></td>
          <td><div class=value>���</div></td>
          <td><div class=value>������</div></td>
          <td><div class=value>ͼʾ</div></td>
        </tr>
        <%Dim N
        For N = 1 to DEF_UserLevelNum%>
        <tr>
          <td class=tdbox><%=N%>��</td>
          <td class=tdbox><%=DEF_UserLevelString(N)%></td>
          <td class=tdbox><%=DEF_UserLevelPoints(N)%></td>
          <td class=tdbox><img src=../../images/<%=GBL_DefineImage%>lvstar/level<%=N%>.gif height=11 width=110></td>
        </tr><%Next%>
      </table>
      </div>
</ol>
<a name="FullSearchHelp"><div class=title>ȫ�ļ���ʹ�ð���</div></a>

                        <p><b>��������</b></P>
                        <ul>
                        <p>������෽�㣬����Ҫ��Ҫ�������ѯ���ݲ���һ�»س��������ɿ�ʼ��ѯ��</P>
                        <p>�����Ͻ����棬�Բ�ѯҪ��һ�ֲ�����磺�ԡ�Ӧ�á��������͡�ʹ�á���<br>
                        ����������ֲ�ͬ�Ľ�������������ʱ�����������ò�ͬ�Ĺؼ��ʣ����ǣ���<br>
                        �������ִ�Сд�����������ASP���͡�aSP����û������ġ�
                        </ul>
                        
                        <p><b>AND��ʹ��</b></P>
                        <ul>
                        <p>�ڼ���ʱ����Ҫʹ��"and"������ϵͳ�Զ����ڹؼ���֮���Զ����"AND"���ṩ<br>
                        ������ȫ����ѯ�����ļ�¼�������������С����������Χ��ֻ����������<br>
                        �ؼ��ʡ�
                        <p>���磺������Ȱ����С�SQL�����С���䡱��ֻ����롰SQL ��䡱���ɣ�Ȼ��<br>
                        ����������ť�����������롰SQL and ��䡱��<br>
                        </P>
                        </ul>
                        <p><b>�����С������Χ?</b></P>
                        <ul>
                          <p>��ʱ��ѯ��õ�����Ľ����Ϊ�õ���ʵ�õ����ϣ�����Ҫ��һ����С��<br>
                          ѯ����ֻҪ�������Ĺؼ���ɸѡ��ѯ���������ϣ�������С������Χ��</P>
                        </ul>
                        <p><b>��OR����AND���Ƿ���Ч?</b></P>
                        <ul>
                          <p>�ڼ����мȲ�ʹ�á�AND��Ҳ��ʹ�á�OR�������ڲ�֧�֡�OR������,���Լ�<br>
                          ��ʱ�޷����ܡ����߰�������A�����߰�������B���ļ�¼���磺��Ҫ��ѯ��<br>
                          SQL����Oracle�����ͱ�������β�ѯ�ֱ��ѯ��SQL���͡�Oracle����</P>
                        </ul>
                        <p><b>���Դ���</b></P>
                        <ul>
                          <p>ͨ�������ǻ����<EM>���ҡ�</EM>��<EM>���ġ�</EM>��̫����������̫��������ַ����Լ���<br>
                          �ֺ͵���ĸ��һЩ�ַ���</P>
                        </ul>
                        <p><b>Ϊʲô�еĲ�ѯ�����¼������ȷ��</b></P>
                        <ul>
                          <p>Ϊ���ṩ��ѯ�ٶȣ�ֻ�ṩ���ϼ�¼��ǰ��һ���ֹ���ѯ�����ܲ�ѯ�����<br>
                          ��¼Ҳֻ�����ƴ���</P>
                        </ul>
                      <a name=003><div class=title>�������</div></a>
                      <br>���������:<br>
                        <ul>
                          <li>Internet Explorer 6.0+��Firefox��Chrome��Safari��Opera</li>
                        </ul>
                        �������:<br>
                        <ul><li>
                          Internet Explorer 10.0��Firefox��Chrome��Safari��Opera</li>
                          <br>
                          Ϊ�˸��õر���������˽��Ϣ�������������ȫ����ȫ����ʹ�ñ�վ���ܣ�<br>
                          ������ʹ�ø���׼��������°汾��</li>
                        </ul>
                        <p>
                        ������Ļ�ֱ���
                        <ul>
                          1440 x 900 ������
                        </ul>
                        <p>�Զ�ת��ͼ������֧�ֵ�ͼ���ʽ<p>
                        <ul>
                         jpg gif jpeg jpe png bmp psd tif sgi tga iff pcx dcx pbm pgm ppm<br>
                         pnm miff xbm xpm ico icl emf hru jif prc wrl wbmp
                        </ul>
                        <a name=code></a><b><font color=red class=redfont>ʲô����֤�룿</font></b>
                        <ol><li>��֤��Ҫ���û�������������֤ʱ����Ҫ����ҳ������ͼƬ��ʾ���ַ�����
				<li>��������Ϊ�˱��ⲻ���û�ʹ�ó������̳���й�ˮ�ͷ���������棬���ṩ�Ĺ��ܣ�ʹ�û��õ�����ȫ�ķ���
				<li>�û����ü��丽���룬������ֻ�Ե��η�����Ч�������ظ�ʹ�á�
						</ol>
                        <p>
                        <a name=lmt></a><b><a name=004><div class=title>�����������</div></a></b>
                        <ul>
                        ��ĳЩԭ����̳���ܻ��һЩ���漰�������ݽ��������������<p>
                        <b>�������������¼��������</b><p>
                        <ol>
                        <li>ֻ�е�¼�û����ܷ��ʣ�����<a href=../<%=DEF_RegisterFile%>>ע��</a>��Ϊ��̳�û�������<a href=../login.asp>��¼</a>��������
                        <li>ֻ��<%=DEF_PointsName(8)%>���Ͽ��ţ�ֻ��<%=DEF_PointsName(8)%>��������İ��棬һ����<%=DEF_PointsName(8)%>����ר�ð���
                        <li>ֻ��<%=DEF_PointsName(5)%>���ţ���̳�û�����ͨ��<%=DEF_PointsName(5)%>���û���֤��Ҫ������Ա���ָ��
                        <li>������̳���ڽ���֮ǰ��ÿ���û�����������Ӧ������(�еĻ���������֤��)�������ɹ���Ա�ṩ
                        <li>��̳����п�����̳���������ο�ֱ�ӷ�����ظ�����
                        <li>��ʱ�ر���̳�����������ʱ���ԣ�����ʾ��ʱ����ʹر�״̬��
                        <li>�����д�������Ա��˲��ܲ鿴������������Ҫ������Ա���(����)ͨ�����ܲ鿴
                        <li>����������ƣ���Ҫ�û������ض�״̬���ܷ���
                        </ol>
                        
                        <br><b>���������������������¼��������</b><p>
                        <ol>
                        <li>�鿴������Ҫһ��<%=DEF_PointsName(0)%>��<%=DEF_PointsName(0)%>һ���ɷ������ӻ�����;����ã�ּ��Ϊ��̳����������鿴ֻ��Ҫ�ﵽһ��<%=DEF_PointsName(0)%>����������<%=DEF_PointsName(0)%>ֵ
                        <li>�鿴������Ҫһ��<%=DEF_PointsName(4)%>��<%=DEF_PointsName(4)%>���ݵ�¼����û���������ʱ�������㣬һ��������<%=DEF_PointsName(4)%>1���Ǻ����û�������̳ʱ��ı�׼���鿴���Ӻ󲢲�����<%=DEF_PointsName(4)%>ֵ
                        <li>���������Ҫ����<%=DEF_PointsName(0)%>���鿴���������ӣ����ȱ����ǵ�¼����û�������Ҫ����ʾ��<%=DEF_PointsName(0)%>ֵ���������ϵͳ���������Ӧ<%=DEF_PointsName(0)%>����ת�Ƹ������ˡ�����󣬽��������������������ݡ�
                        <li>������<%=DEF_PointsName(8)%>���ܲ鿴���������ӽ�����˰�<%=DEF_PointsName(8)%>��<%=DEF_PointsName(6)%>���ϵĹ���Ա��Ա�鿴��
                        <li>��<%=DEF_PointsName(8)%>���ܲ鿴����������ֻ��<%=DEF_PointsName(8)%>����Ȩ�޵��û����ܲ鿴��
                        <li>��<%=DEF_PointsName(5)%>���ܲ鿴���鿴�������ӱ����ȳ�Ϊ��̳��<%=DEF_PointsName(5)%>����̳�û�����ͨ����֤�û����û���֤��Ҫ������Ա���ָ����
                        </ol>
                        <p>
                        <a name=lmt></a><a name=005><div class=title>��̳�����û����<%=DEF_PointsName(9)%>һ����</div></a>
                        <br>
                        <ol>
                        <li>��ͨ�û���ӵ�������̳���������ӺͶ���Ϣ�Ȼ�����Ȩ��</li>
                        <li>����ʽ�û���ֻӵ�������̳��Ȩ�ޣ�����Ȩ�޷������Ӻͷ��Ͷ���Ϣ��ͶƱ�Ȳ���</li>
                        <li><%=DEF_PointsName(5)%>��ӵ�н���ĳЩ��֤��������Ȩ��</li>
                        <li><%=DEF_PointsName(8)%>��ӵ�ж������ΰ���ά��Ȩ��</li>
                        <li><%=DEF_PointsName(6)%>��ӵ��ȫ������ά�����̶ܹ������Ȩ�ޣ���������չȨ��</li>
                        <li><%=DEF_PointsName(9)%>һ����
                        	<ul>
				<%
				For N = 1 to DEF_UserOfficerNum
					Response.Write "<li>���" & N & "��" & DEF_UserOfficerString(N) & "</li>" & VbCrLf
				Next
				%>
                        	</ul>
                        </li>
                        </ol>

                        <p>
                        <a name=lmt></a><a name=006><div class=title>��̳����ͼ��</div></a>
                        <br>
                        <ul>
                        <div class=b_new><br>�������İ���<br><br></div>
                        <br>
                        <div class=b_none><br>�������İ���<br><br></div>
                        </ul>

                        <p>
                        <a name=lmt></a><a name=007><div class=title>��̳����ͼ��</div></a>
                        <br>
                        <ul>
                        <img src=../../images/<%=GBL_DefineImage%>state/alltop.gif align=absmiddle>
                        �̶ܹ������ӣ��κΰ���ɼ���
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/parttop.gif align=absmiddle>
                        ���̶������ӣ������̶ܹ���֮�¹̶���֮�ϣ����ڵ�ǰ������Ч
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/intop.gif align=absmiddle>
                        �̶������ӣ���Զ�̶��ڰ������ҳ����λ��
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/vt.gif align=absmiddle>
                        ͶƱ��
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/tpcnew.gif align=absmiddle>
                        ������
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/hot.gif align=absmiddle>
                        ��������
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/tpc.gif align=absmiddle>
                        ��ͨ����
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>state/lock.gif align=absmiddle>
                        ����������
                        <p>
                        </ul>
                        

                        <p>
                        <a name=lmt></a><a name=008><div class=title>��̳���ܼ�������ͼ��</div></a>
                        <br>
                        <ul>
                        <img src=../../images/<%=GBL_DefineImage%>home.gif align=absmiddle>
                        �鿴�������ߵ���վ��ҳ
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>mail.gif align=absmiddle>
                        ���������߷����ʼ�(ʹ���ʼ��������)
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>re.gif align=absmiddle>
                        ���ô����ӵĲ������ݽ��лظ�
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>message.gif align=absmiddle>
                        ���������߷�����̳����Ϣ
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>edit.gif align=absmiddle>
                        �༭��̳��������
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>ts.gif align=absmiddle>
                        �Զ��Ű��������໻�� ������Ա�����������������ӣ��������
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>friend.gif align=absmiddle>
                        �Ӵ��û���Ϊ�ҵ���̳����
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>collect.gif align=absmiddle>
                        �ղ���̳����
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>del.gif align=absmiddle>
                        ɾ�������ӻ�Ͷ�����վ
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>jh.gif align=absmiddle>
                        ����һ�������ȡ�����������ͷ����û�
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>ti.gif align=absmiddle>
                        ��ȡ���⵽���ڰ��������λ��
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>repair.gif align=absmiddle>
                        �Զ��޸����˵�����
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>move.gif align=absmiddle>
                        ת�ƴ����⵽��������̳����
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>maketop.gif align=absmiddle>
                        ���������Ϊ�̶����ӻ�ȡ���̶�
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>makeparttop.gif align=absmiddle>
                        ���������Ϊ���̶����ӻ�ȡ�����̶�
                        <p>
                        <img src=../../images/<%=GBL_DefineImage%>makealltop.gif align=absmiddle>
                        ���������Ϊ�̶ܹ����ӻ�ȡ���̶ܹ�
                        <p>
                        </U>
                      </td>
                    </tr>
                       

                  </table>

<%End Sub%>