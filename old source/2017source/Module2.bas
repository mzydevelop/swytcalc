Attribute VB_Name = "Module2"
Public k As Integer
Public tag As Integer
Public Sub sjk()
Select Case xh
Case 1
xm = "�㽭��ҵ��ѧ"
If ag < 5 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = 100 - 5 * (10 - ag) + 2 * bg
xkf = cxk
End If
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zjut.edu.cn/zsnews/html/n1753.html"
bm = "http://zs.zjut.edu.cn/swytyun/apply/main.jsp"
btime = "���ϱ���ʱ��Ϊ2017��2��20������9:00��3��2������16:00���ʼĲ��Ͻ��մӱ���ʱ����3��2�ս�ֹ���Ե����ʴ�Ϊ׼��"
Case 2
xm = "�㽭ʦ����ѧ"
If ag < 3 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjnu.edu.cn/2017/0123/c101a118137/page.htm"
bm = "http://zsb.zjnu.edu.cn/apply/main.jsp"
btime = "����ʱ�䣺2017��2��22����3��4��"
Case 3
xm = "������ѧ"
If ag < 3 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = 100 - 5 * (10 - ag)
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zsb.nbu.edu.cn/Article/Index/219"
bm = "http://zsb.nbu.edu.cn/Students/Login"
btime = "���ϱ����������ϴ�ʱ�䣺2017��2��25������10:00��3��10������15:00��"
Case 4
xm = "���ݵ��ӿƼ���ѧ"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk / 1.5
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zhaosheng.hdu.edu.cn/art.php?aid=1255"
bm = "http://swyt.hdu.edu.cn/"
btime = "���ϱ�����2017��2��8����3��8�գ�"
Case 5
xm = "�㽭���̴�ѧ"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaoban.zjsu.edu.cn/html/2017-1/201712390751.htm"
bm = "http://zhaoban.zjsu.edu.cn/swyt/index.asp"
btime = "���ϱ���ʱ�䣺2017��2��19����3��1������16:00��"
Case 6
xm = "�㽭����ѧ"
If ag * 15 + bg * 9 + cg * 3 < 100 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk
End If
examb = 0.5
wz = "http://zs.zstu.edu.cn/?p=read&aid=1759"
bm = "http://120.55.84.14/zjlg/Login.aspx"
btime = "����ʱ��Ϊ��������2017��3��1�ա�"
Case 7
xm = "����ҽ�ƴ�ѧ"
If ag < 5 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = 100 - 5 * (10 - ag)
xkf = cxk
End If
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaosheng.wmu.edu.cn/Art/Art_1/Art_1_1329.aspx"
bm = "http://swyt.wmu.edu.cn/apply/main.jsp"
btime = "����������2017��2��10������8:30-2��26������16:00��"
Case 8
xm = "�㽭�����ѧ"
If ag < 2 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjou.edu.cn/info/1047/3021.htm"
bm = "http://swyt.zjou.edu.cn/user.asp"
btime = "���ϱ���ʱ�䣺2017��2��19�ա�2017��3��9�գ�"
Case 9
xm = "�㽭ũ�ִ�ѧ"
If ag < 3 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zafu.edu.cn/info/1016/2644.htm"
bm = "http://swyt.zafu.edu.cn/index.php/Login/index.html"
btime = "���ϱ���ʱ�䣺2017��2��21����3��6������4��00��"
Case 10
xm = "�㽭��ҽҩ��ѧ"
If ag < 5 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
End If
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.zcmu.edu.cn/news_show.asp?id=927"
bm = "http://swytbm.zcmu.edu.cn/stu_login.aspx"
btime = "���ϱ�������������ʼ�ʱ�䣺2017��3��1��9��00-2017��3��15��16��00ֹ"
Case 11
xm = "�й�������ѧ"
If ag < 2 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * dg
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=1763"
bm = "http://swytzs.cjlu.edu.cn:8080/SWYTJL/AppMain.jsp"
btime = "���ϱ���ʱ�䣺2017��3��3��9:00��3��22��16:00��"
Case 12
xm = "�㽭����ѧԺ"
cxk = ag * 10 + bg * 6 + cg * 3 + 1 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zswnew.zwu.edu.cn/bkszsw/72/ca/c4804a94922/page.htm"
bm = "https://swyt.zwu.edu.cn/renderLogin.do;JSESSIONID=d08cd28b-bbfc-49b7-88b8-3128d07f9fa3"
btime = "���ϱ���ʱ�䣺2017��3��1�ա�3��20�ա�"
Case 13
xm = "�㽭�Ƽ�ѧԺ(��ͨ�ࣩ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsb.zust.edu.cn/IndexPage!wzxq.htm?id=3AA16DD86E67476CAE5820141547B6D6"
bm = "http://zsb.zust.edu.cn/IndexPage!login.htm"
btime = "����ʱ�䣺2017��2��7��-3��6�ա�"
Case 14
xm = "�㽭�ƾ���ѧ"
If (ag * 10 + bg * 5 + cg * 2) < 70 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = ag * 10 + bg * 5 + cg * 2
xkf = cxk
End If
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zufe.edu.cn/info/1002/1623.htm"
bm = "http://swyt.zufe.edu.cn/apply/main.jsp"
btime = "���ϱ�����2017��2��20������9:00��3��1������16:00��"
Case 15
xm = "����ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://admission.zjxu.edu.cn/news/7080/read.shtml"
bm = "http://210.33.29.129/"
btime = "2017��2��20�ա�3��15��"
Case 16
xm = "����ʦ����ѧ"
cxk = ag * 10 + bg * 5 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://bkzs.hznu.edu.cn/Details/20170125/410924201701250924378326.html"
bm = "http://hzsf.vnet1000.com"
btime = "2017��2��1��-3��3��"
Case 17
xm = "����ʦ��ѧԺ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.hutc.zj.cn/2014/Item/Show.asp?m=1&d=2688"
bm = "http://swyt.zjhu.edu.cn/apply/main.jsp"
btime = "2017��2��22��-3��5�ա�"
Case 18
xm = "��������ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.usx.edu.cn/info/1010/2541.htm"
bm = "https://zsbm.usx.edu.cn/w3/index.html"
btime = "2017��2��22��-3��5�ա�"
Case 19
xm = "̨��ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.zsjy.tzc.edu.cn/articles/1378.html"
bm = "http://swyt.tzc.edu.cn/apply/main.jsp"
btime = "2017��2��16��-3��1�ա�"
Case 20
xm = "���ݴ�ѧ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = 5 * cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.wzu.edu.cn/Art/Art_5/Art_5_531.aspx"
bm = "http://swyt.wzu.edu.cn:8088/apply/main.jsp"
btime = "2017��2��20������9:00��3��2������16:00��"
Case 21
xm = "�㽭�����ѧԺ"
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://zsjy.zisu.edu.cn/zhaosheng/article.asp?id=1058&?typeid=2"
bm = "http://swyt.zisu.edu.cn/"
btime = "���ϱ���ʱ�䣺2017��2��10����3��9�գ�"
Case 22
xm = "��������ѧԺ"
tag = 0
For k = 1 To 3
  If xkb(k) = "A" Then
  tag = tag + 1
  End If
Next k
cxk = ag * 10 + bg * 5 + cg * 3 + tag * 5
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.nbut.edu.cn/index.php?s=/Home/Article/detail/id/692.html"
bm = "http://zs.nbut.edu.cn/index.php?s=/Home/Index/login.html"
btime = "���ϱ���ʱ�䣺2017��2��20�ա�3��3�գ�"
Case 23
xm = "����ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://xsc.qzu.zj.cn/zs/info_1193.aspx"
bm = "http://xsc.qzu.zj.cn:8081/apply/main.jsp"
btime = "���ϱ���ʱ�䣺2017��2��17����3��2�գ�"
Case 24
xm = "�㽭ˮ��ˮ��ѧԺ"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zjc.zjweu.edu.cn/zhaosheng/d3/35/c1620a54069/page.htm"
bm = "http://swyt.zjweu.edu.cn/"
btime = "���ϱ�����2017��2��22��-3��8�գ�"
Case 25
xm = "��ˮѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4 + 1 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.lsu.edu.cn/2017/0124/c590a210179/page.htm"
bm = "http://swyt.lsu.edu.cn/apply/main.jsp"
btime = "2017��2��12�գ�ũ������ʮ������2��25�գ����¶�ʮ�ţ�"
Case 26
xm = "���ݿ϶���ѧ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = 5 * cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.wku.edu.cn/zh-hans/2017/01/wenzhoukenendaxue2017niansanweiyiti-zonghepingjiazhaoshengzhangcheng/"
bm = "http://121.40.33.186:810/"
btime = "���ϱ���ʱ�䣺2017��3��6��9:00��2017��3��27��17:00��"
Case 27
xm = "����ŵ������ѧ"
If ag < 7 Then
MsgBox "δ���ϸ�У��ͱ�������"
Else
cxk = 100 - 5 * (10 - ag)
xkf = cxk
End If
xb = 0.1
zb = 0.3
examb = 0.6
wz = "http://www.nottingham.edu.cn/cn/study/undergraduate/policy-and-requirements/2017-direct-entry.aspx"
bm = "http://125.111.163.252:8080/apply/main.jsp"
btime = "2017��3��13��09��00-5��4��16��00"
Case 28
xm = "�㽭��ѧ����ѧԺ"
tag = 0
For k = 1 To 3
  If xkb(k) = "A" Then
  tag = tag + 1
  End If
Next k
cxk = ag * 10 + bg * 5 + tag * 5
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zucc.edu.cn/baokao_show.aspx?id=887"
bm = "http://swyt.zucc.edu.cn/apply/main.jsp"
btime = "���ϱ���ʱ�䣺2017��2��24������9:00��3��6������16:00��"
Case 29
xm = "�㽭��ѧ������ѧԺ"
tag = 0
For i = 1 To 3
  If xkb(i) = "A" Then
  tag = tag + 1
  End If
Next i
cxk = ag * 10 + bg * 5 + cg * 2 + tag * 5
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.nit.net.cn/info/1022/2287.htm"
bm = "http://swyt.nit.zju.edu.cn/apply/main.jsp"
btime = "2017��2��24������9:00��3��12������17:00;"
Case 30
xm = "�㽭����ѧԺ���㽭���˴�ѧ��"
cxk = ag * 10 + bg * 7 + cg * 3 + dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjsru.edu.cn/Cms/CmsDetail/154?20160701ba0b5679"
bm = "http://swyt.zjsru.edu.cn/apply/main.jsp"
btime = "���ϱ���ʱ�䣺2017��2��17�� - 3��7�ա�"
Case 31
xm = "�㽭Խ�������ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zyufl.edu.cn/xsrx_1.aspx?newsid=557"
bm = "http://zs.yxc.cn/3/index.aspx"
btime = "���ϱ���ʱ�䣺2017��2��16����3��16�գ����ڲ�������"
Case 32
xm = "�������ӥѧԺ"
cxk = ag * 12 + bg * 8 + cg * 4 + 2 * dg
If cxk > 100 Then
cxk = 100
End If
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.dhyedu.com/zsb/show.htm?id=11175"
bm = "http://zsb.nbdhyu.edu.cn/"
btime = "���ϱ���ʱ�䣺2017��2��10�ա�3��20�ա�"
Case 33
xm = "����ҽ�ƴ�ѧ�ʼ�ѧԺ"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://rjxy.wmu.edu.cn/zhaosheng/view.jsp?id0=&id=z0iyato2yn"
bm = "http://swyt.rjxy.wmu.edu.cn/apply/main.jsp"
btime = "2017��2��10��-2��26�ա�"
Case 34
xm = "�㽭��ҽҩ��ѧ����ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.zcmu.edu.cn/news_show.asp?id=928"
bm = "http://swytbm.zcmu.edu.cn/stu_login.aspx"
btime = "2017��3��1��9��00-2017��3��15��16��00ֹ"
Case 35
xm = "�й�������ѧ�ִ��Ƽ�ѧԺ"
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * (10 - ag - bg - cg)
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=1762"
bm = "http://swytzs.cjlu.edu.cn:8080/SWYTXDKJ/AppMain.jsp"
btime = "���ϱ���ʱ�䣺2017��3��3��9:00��3��22��16:00��"
Case 36
xm = "����ʦ����ѧǮ��ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 5
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://qjzs.hznu.edu.cn/shownews.asp?id=758"
bm = "http://qjzsbm.hznu.edu.cn/zsbm_show/login.asp"
btime = "����ʱ�䣺2017��2��11����3��5�ա�"
Case 37
xm = "������ѧԺ"
cxk = ag * 12 + bg * 8 + cg * 5 + 2 * dg
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zsw.wzbc.edu.cn/Art/Art_35/Art_35_4305.aspx"
bm = "http://3w1t.wzbc.edu.cn/apply/main.jsp"
btime = "���ϱ�����2017��2��27�ա�3��26��"
Case 38
xm = "ͬ�ô�ѧ�㽭ѧԺ"
cxk = ag * 10 + bg * 5 + cg * 3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://www.tjzj.edu.cn/zhaosheng.php?pid=745&cid=763&id=341212"
bm = "http://122.225.19.18:8081/apply/main.jsp"
btime = "2017��2��25������9��00��3��6������16��00��"
Case 39
xm = "�Ϻ��ƾ���ѧ�㽭ѧԺ"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.shufe-zj.edu.cn/index.aspx?lanmuid=63&sublanmuid=65&id=6315"
bm = "http://swyt.shufe-zj.edu.cn/apply/main.jsp"
btime = "����ʱ�䣺2017��2��21��9:00��3��3��15:00"
Case Else
MsgBox "ERROR���Ҳ���������", vbExclamation, "ϵͳ��Ϣ"
End Select
End Sub

