Attribute VB_Name = "Module2"
Public k As Integer
Public tag As Integer
Public Sub sjk()
Select Case xh
Case 1
xm = "�㽭��ҵ��ѧ"
cxk = 100 - 5 * (10 - ag) + 2 * bg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zjut.edu.cn/zsnews/html/n1908.html"
bm = "http://zs.zjut.edu.cn/swytyun/apply/main.jsp"
btime = "���ϱ���ʱ��Ϊ2018��3��5������10:00��3��19������16:00��"
Case 2
xm = "�㽭ʦ����ѧ"
cxk = ag * 10 + bg * 8 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjnu.edu.cn/2018/0206/c6828a232320/page.htm"
bm = "http://zsb.zjnu.edu.cn/apply/main.jsp"
btime = "2018��3��9����3��24��"
Case 3
xm = "������ѧ"
cxk = 100 - 5 * (10 - ag)
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.nbu.edu.cn/Article/Index/261"
bm = "http://zsb.nbu.edu.cn/Students/Login"
btime = "2018��3��1������10:00��3��13������15:00��"
Case 4
xm = "���ݵ��ӿƼ���ѧ"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk / 1.5
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zhaosheng.hdu.edu.cn/art.php?aid=1378"
bm = "http://swyt.hdu.edu.cn/"
btime = "���ϱ�����2018��2��22����3��20�գ�"
Case 5
xm = "�㽭���̴�ѧ"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaoban.zjsu.edu.cn/View-1518.html"
bm = "http://zjgsu.swytbm.com/#/login"
btime = "2018��3��4����3��23������16:00��"
Case 6
xm = "�㽭����ѧ"
cxk = ag * 15 + bg * 12 + cg * 6
xkf = cxk
examb = 0.5
wz = "http://zs.zstu.edu.cn/info/1004/2855.htm"
bm = "http://120.55.84.14:8080/zjlg/Login.aspx"
btime = "����ʱ��Ϊ��������2018��3��16��16ʱ��"
Case 7
xm = "����ҽ�ƴ�ѧ"
cxk = 10*ag+5*bg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaosheng.wmu.edu.cn/Art/Art_1/Art_1_3078.aspx"
bm = "http://swyt.wmu.edu.cn/apply/main.jsp"
btime = "2018��2��26������8:30��3��18������16:00��"
Case 8
xm = "�㽭�����ѧ"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjou.edu.cn/info/1047/3600.htm"
bm = "http://swyt.zjou.edu.cn/user.asp"
btime = "���ϱ���ʱ�䣺2018��3��1�ա�2018��3��15�գ�"
Case 9
xm = "�㽭ũ�ִ�ѧ"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk/1.5
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zafu.edu.cn/info/1014/3026.htm"
bm = "http://swyt.zafu.edu.cn/index.php/Login/index.html"
btime = "2018��3��5����3��19������4��00��"
Case 10
xm = "�㽭��ҽҩ��ѧ"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://swytbm.zcmu.edu.cn/swytyun/apply/main.jsp"
bm = "http://swytbm.zcmu.edu.cn/swytyun/apply/main.jsp"
btime = "2018��3��4��9��00-2018��3��17��16��00ֹ"
Case 11
xm = "�й�������ѧ"
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=2133"
bm = "http://swytzs.cjlu.edu.cn:8080/SWYTJL/AppMain.jsp"
btime = "���ϱ���ʱ�䣺2018��3��5��9:00��3��25��16:00��"
Case 12
xm = "�㽭����ѧԺ"
cxk = ag * 12 + bg * 9 + cg * 6 + 4 * dg
If cxk>100 Then
cxk=100
End If
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zsw.zwu.edu.cn/bkszsw/97/65/c4804a104293/page.htm"
bm = "https://swyt.zwu.edu.cn/renderLogin.do;JSESSIONID=d08cd28b-bbfc-49b7-88b8-3128d07f9fa3"
btime = "2018��3��8�ա�3��20�ա�"
Case 13
xm = "�㽭�Ƽ�ѧԺ(��ͨ�ࣩ"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsb.zust.edu.cn/IndexPage!wzxq.htm?id=7790C245202746EA9F171220742E8474"
bm = "http://zsb.zust.edu.cn/IndexPage!login.htm"
btime = "����ʱ�䣺2018��2��10��-3��25�ա�"
Case 14
xm = "�㽭�ƾ���ѧ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk
examb = 0.5
wz = "http://zs.zufe.edu.cn/info/1002/1803.htm"
bm = "http://swyt.zufe.edu.cn/apply/main.jsp"
btime = "2018��2��26������9:00��3��10������16:00��"
Case 15
xm = "����ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.zjxu.edu.cn/news/7107/read.shtml"
bm = "http://210.33.29.129/"
btime = "����������2018��3��3����3��16�գ�"
Case 16
xm = "����ʦ����ѧ���ǵ�������"
cxk = ag * 10 + bg * 5 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.3
examb = 0.6
wz = "http://bkzs.hznu.edu.cn/Details/20180207/410335201802070335635566.html"
bm = "http://hzsf.vnet1000.com"
btime = "����ʱ�䣺2018��3��1�ա�3��23�ա�"
Case 17
xm = "����ʦ��ѧԺ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.zjhu.edu.cn/2014/Item/Show.asp?m=1&d=2785"
bm = "http://swyt.zjhu.edu.cn/apply/main.jsp"
btime = "����ʱ��Ϊ2018��2��22��-3��7�ա�"
Case 18
xm = "��������ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.usx.edu.cn/info/1020/3321.htm"
bm = "https://zsbm.usx.edu.cn/w3/index.html"
btime = "2��20������3��18�գ���"
Case 19
xm = "̨��ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.zsjy.tzc.edu.cn/articles/1583.html"
bm = "http://swyt.tzc.edu.cn/apply/main.jsp"
btime = "2018��3��1��-3��14�ա�"
Case 20
xm = "���ݴ�ѧ"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = 5 * cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.wzu.edu.cn/info/1010/1702.htm"
bm = "http://swyt.wzu.edu.cn:8088/apply/main.jsp"
btime = "2018��3��4������9:00��3��20������16:00��"
Case 21
xm = "�㽭�����ѧԺ"
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://swyt.zisu.edu.cn/newsShow.aspx?newsCate=9ZaCYPKsse9I1m4n_et1fQ==&newsID=59"
bm = "http://swyt.zisu.edu.cn/"
btime = "���ϱ���ʱ�䣺2018��2��12����3��9�գ�"
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
wz = "http://zs.nbut.edu.cn/index.php?s=/Home/Article/detail/id/896.html"
bm = "http://zs.nbut.edu.cn/index.php?s=/Home/Index/login.html"
btime = "2018��3��1�ա�3��16�գ�"
Case 23
xm = "����ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://xsc.qzu.zj.cn/zs/info_1261.aspx"
bm = "http://xsc.qzu.zj.cn:8081/apply/main.jsp"
btime = "���ϱ���ʱ�䣺2018��3��5����3��18�գ���"
Case 24
xm = "�㽭ˮ��ˮ��ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4 + 1 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zjc.zjweu.edu.cn/zhaosheng/f6/2b/c1620a63019/page.htm"
bm = "http://swyt.zjweu.edu.cn/"
btime = "���ϱ�����2018��3��5��-3��25�գ�"
Case 25
xm = "��ˮѧԺ"
cxk = ag * 10 + bg * 7 + cg * 4 + 1 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.lsu.edu.cn/2018/0207/c619a253907/page.htm"
bm = "http://47.97.34.107/apply/main.jsp"
btime = "2018��3��4�գ�ũ������ʮ�ߣ���3��18�գ�ũ�����³�����"
Case 26
xm = "���ݿ϶���ѧ"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk/1.5
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.wku.edu.cn/zh-hans/2018/02/20183in1admissions/"
bm = "http://121.40.33.186:810/"
btime = "���ϱ���ʱ�䣺2018 ��3��5�� 9:00��2018��3��26�� 17:00��"
Case 27
xm = "����ŵ������ѧ"
cxk = 100 - 5 * (10 - ag)
xkf = cxk
xb = 0.1
zb = 0.3
examb = 0.6
wz = "https://www.nottingham.edu.cn/cn/study/undergraduate/policy-and-requirements/2018-direct-entry.aspx"
bm = "http://triunity.nottingham.edu.cn/apply/main.jsp"
btime = "���ϱ���ʱ�䣺2018��3��15������9:00��5��9������16:00"
Case 28
xm = "�㽭��ѧ����ѧԺ"
tag = 0
For k = 1 To 3
  If xkb(k) = "A" Then
  tag = tag + 1
  End If
Next k
cxk = ag * 10 + bg * 5 + tag * 5 +cg*2
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zucc.edu.cn/baokao_show.aspx?id=970"
bm = "http://swyt.zucc.edu.cn/apply/main.jsp"
btime = "���ϱ���ʱ�䣺2018��3��1����3��11�գ�"
Case 29
xm = "�㽭��ѧ������ѧԺ"
tag = 0
For i = 1 To 3
  If xkb(i) = "A" Then
  tag = tag + 1
  End If
Next i
cxk = ag * 10 + bg * 8 + cg * 4 + tag * 5
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.nit.net.cn/info/1022/2629.htm"
bm = "http://swyt.nit.zju.edu.cn/apply/main.jsp"
btime = "2018��3��4������9:00��29������17:00;"
Case 30
xm = "�㽭����ѧԺ���㽭���˴�ѧ��"
cxk = ag * 10 + bg * 7 + cg * 3 + dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjsru.edu.cn/Cms/CmsDetail/182"
bm = "https://swyt.zjsru.edu.cn/apply/main.jsp"
btime = "2018��3��2�� - 3��19�ա�"
Case 31
xm = "�㽭Խ�������ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.yxc.cn/xsrx_1.aspx?newsid=619"
bm = "http://zs.yxc.cn/3/index.aspx"
btime = "2018��2��26����3��18�գ����ڲ�������"
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
wz = "http://www.nbdhyu.edu.cn/zsb/show.htm?id=14510"
bm = "http://zsb.nbdhyu.edu.cn/"
btime = "2018��2��23�ա�3��22�ա�"
Case 33
xm = "����ҽ�ƴ�ѧ�ʼ�ѧԺ"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaosheng.wmu.edu.cn/Art/Art_1/Art_1_3080.aspx"
bm = "http://swyt.rjxy.wmu.edu.cn/apply/main.jsp"
btime = "2018��2��26������8:30-3��18������16:00��"
Case 34
xm = "�㽭��ҽҩ��ѧ����ѧԺ"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://swytbm.zcmu.edu.cn/swytyunbjxy/apply/main.jsp"
bm = "http://swytbm.zcmu.edu.cn/swytyunbjxy/apply/main.jsp"
btime = "2018��3��4��9��00-2018��3��17��16��00ֹ"
Case 35
xm = "�й�������ѧ�ִ��Ƽ�ѧԺ"
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=2144"
bm = "http://swytxk.cjlu.edu.cn/apply/main.jsp"
btime = "2018��3��5��9:00��3��25��16:00��"
Case 36
xm = "����ʦ����ѧǮ��ѧԺ"
cxk = ag * 10 + bg * 7 + cg * 5
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://qjzs.hznu.edu.cn/shownews.asp?id=816"
bm = "http://qjzsbm.hznu.edu.cn/zsbm_show/login.asp"
btime = "����ʱ�䣺2018��3��3����3��25��"
Case 37
xm = "������ѧԺ"
cxk = ag * 15 + bg * 11 + cg * 7 + 3 * dg
If cxk > 100 Then
cxk = 100
Else
cxk = cxk
End If
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zsw.wzbc.edu.cn/Art/Art_70/Art_70_7612.aspx"
bm = "http://3w1t.wzbc.edu.cn/apply/main.jsp"
btime = "2018��3��5�ա�3��31��"
Case 38
xm = "ͬ�ô�ѧ�㽭ѧԺ"
cxk = ag * 10 + bg * 5 + cg * 3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://www.tjzj.edu.cn/zhaosheng.php?pid=745&cid=763&id=343313"
bm = "http://122.225.19.18:8081/apply/main.jsp"
btime = "2018��3��13������9��00��3��22������16��00��"
Case 39
xm = "�Ϻ��ƾ���ѧ�㽭ѧԺ����ҳ�޷����ʣ�"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.shufe-zj.edu.cn/index.aspx?lanmuid=63&sublanmuid=65&id=6315"
bm = "http://swyt.shufe-zj.edu.cn/apply/main.jsp"
btime = "����ʱ�䣺2017��2��21��9:00��3��3��15:00"
Case 40
xm = "����ҽѧԺ"
cxk = ag * 10 + bg * 7 +cg*3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.hmc.edu.cn/art/2018/2/8/art_2399_132743.html"
bm = ""
btime = "���ϱ���ʱ��Ϊ2018��3��12������9:00��3��22������16:00"
Case 41
xm = "�㽭��ҵ��ѧ֮��ѧԺ"
cxk = ag * 10 + bg * 5 +cg*2
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zjc.zjut.edu.cn/archive/1059.html"
bm = "http://zs.zjc.zjut.edu.cn"
btime = "2018��3��5������9:00-3��19������16:00��"
Case 42
xm = "������ѧ��ѧ����ѧԺ "
cxk = ag * 10 + bg * 7 +cg*3
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://zs.ndky.edu.cn/prospectuses/7427.jhtml"
bm = "http://swyt.ndky.edu.cn/apply/main.jsp"
btime = "2018��3��1������10:00��3��26������15:00��"
Case 43
xm = "���ݵ��ӿƼ���ѧ��Ϣ����ѧԺ"
cxk = ag * 10 + bg * 5 +cg*2
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://www.hziee.edu.cn/index.php?c=Index&a=news_detail&catid=453&id=2812&web=bkszs"
bm = "http://swyt.hziee.edu.cn/"
btime = "2018��3��5�ա�2018��3��25�գ�"
Case 44
xm = "�㽭�ƾ���ѧ����ѧԺ"
cxk = ag * 10 + bg * 8 +cg*5+dg*2
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zufedfc.edu.cn/index.php?mod=show&mid=3&pid=193&id=2728"
bm = "http://zsbm.zufedfc.edu.cn:8080/apply/main.jsp"
btime = "2018��3��1������9��00��3��22������16��00"
Case 45
xm = "��������ѧԺԪ��ѧԺ"
cxk = ag * 10 + bg * 8 +cg*7+dg*6
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsxx.ypc.edu.cn/info/1002/1424.htm"
bm = "http://zsxx.ypc.edu.cn/"
btime = "���ϱ�����2018��3��1����3��25�գ�"
Case 46
xm = "���ݴ�ѧ걽�ѧԺ"
cxk = ag * 10 + bg * 6 +cg*3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.ojc.zj.cn/Art/Art_383/Art_383_61403.aspx"
bm = "http://ojzs.cnvp.com.cn/apply/main.jsp"
btime = "2018��2��25������9:00��3��15������16:00��"
Case Else
MsgBox "ERROR���Ҳ���������", vbExclamation, "ϵͳ��Ϣ"
End Select
End Sub

