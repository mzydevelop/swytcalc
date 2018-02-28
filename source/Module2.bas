Attribute VB_Name = "Module2"
Public k As Integer
Public tag As Integer
Public Sub sjk()
Select Case xh
Case 1
xm = "浙江工业大学"
cxk = 100 - 5 * (10 - ag) + 2 * bg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zjut.edu.cn/zsnews/html/n1908.html"
bm = "http://zs.zjut.edu.cn/swytyun/apply/main.jsp"
btime = "网上报名时间为2018年3月5日上午10:00至3月19日下午16:00。"
Case 2
xm = "浙江师范大学"
cxk = ag * 10 + bg * 8 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjnu.edu.cn/2018/0206/c6828a232320/page.htm"
bm = "http://zsb.zjnu.edu.cn/apply/main.jsp"
btime = "2018年3月9日至3月24日"
Case 3
xm = "宁波大学"
cxk = 100 - 5 * (10 - ag)
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.nbu.edu.cn/Article/Index/261"
bm = "http://zsb.nbu.edu.cn/Students/Login"
btime = "2018年3月1日上午10:00至3月13日下午15:00；"
Case 4
xm = "杭州电子科技大学"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk / 1.5
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zhaosheng.hdu.edu.cn/art.php?aid=1378"
bm = "http://swyt.hdu.edu.cn/"
btime = "网上报名：2018年2月22日至3月20日；"
Case 5
xm = "浙江工商大学"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaoban.zjsu.edu.cn/View-1518.html"
bm = "http://zjgsu.swytbm.com/#/login"
btime = "2018年3月4日至3月23日下午16:00；"
Case 6
xm = "浙江理工大学"
cxk = ag * 15 + bg * 12 + cg * 6
xkf = cxk
examb = 0.5
wz = "http://zs.zstu.edu.cn/info/1004/2855.htm"
bm = "http://120.55.84.14:8080/zjlg/Login.aspx"
btime = "报名时间为即日起至2018年3月16日16时。"
Case 7
xm = "温州医科大学"
cxk = 10*ag+5*bg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaosheng.wmu.edu.cn/Art/Art_1/Art_1_3078.aspx"
bm = "http://swyt.wmu.edu.cn/apply/main.jsp"
btime = "2018年2月26日上午8:30至3月18日下午16:00；"
Case 8
xm = "浙江海洋大学"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjou.edu.cn/info/1047/3600.htm"
bm = "http://swyt.zjou.edu.cn/user.asp"
btime = "网上报名时间：2018年3月1日—2018年3月15日；"
Case 9
xm = "浙江农林大学"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk/1.5
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zafu.edu.cn/info/1014/3026.htm"
bm = "http://swyt.zafu.edu.cn/index.php/Login/index.html"
btime = "2018年3月5日至3月19日下午4：00。"
Case 10
xm = "浙江中医药大学"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://swytbm.zcmu.edu.cn/swytyun/apply/main.jsp"
bm = "http://swytbm.zcmu.edu.cn/swytyun/apply/main.jsp"
btime = "2018年3月4日9：00-2018年3月17日16：00止"
Case 11
xm = "中国计量大学"
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=2133"
bm = "http://swytzs.cjlu.edu.cn:8080/SWYTJL/AppMain.jsp"
btime = "网上报名时间：2018年3月5日9:00至3月25日16:00；"
Case 12
xm = "浙江万里学院"
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
btime = "2018年3月8日—3月20日。"
Case 13
xm = "浙江科技学院(普通类）"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsb.zust.edu.cn/IndexPage!wzxq.htm?id=7790C245202746EA9F171220742E8474"
bm = "http://zsb.zust.edu.cn/IndexPage!login.htm"
btime = "报名时间：2018年2月10日-3月25日。"
Case 14
xm = "浙江财经大学"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk
examb = 0.5
wz = "http://zs.zufe.edu.cn/info/1002/1803.htm"
bm = "http://swyt.zufe.edu.cn/apply/main.jsp"
btime = "2018年2月26日上午9:00至3月10日下午16:00；"
Case 15
xm = "嘉兴学院"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.zjxu.edu.cn/news/7107/read.shtml"
bm = "http://210.33.29.129/"
btime = "考生报名：2018年3月3日至3月16日；"
Case 16
xm = "杭州师范大学（非电子商务）"
cxk = ag * 10 + bg * 5 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.3
examb = 0.6
wz = "http://bkzs.hznu.edu.cn/Details/20180207/410335201802070335635566.html"
bm = "http://hzsf.vnet1000.com"
btime = "报名时间：2018年3月1日—3月23日。"
Case 17
xm = "湖州师范学院"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.zjhu.edu.cn/2014/Item/Show.asp?m=1&d=2785"
bm = "http://swyt.zjhu.edu.cn/apply/main.jsp"
btime = "报名时间为2018年2月22日-3月7日。"
Case 18
xm = "绍兴文理学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.usx.edu.cn/info/1020/3321.htm"
bm = "https://zsbm.usx.edu.cn/w3/index.html"
btime = "2月20日起至3月18日；。"
Case 19
xm = "台州学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.zsjy.tzc.edu.cn/articles/1583.html"
bm = "http://swyt.tzc.edu.cn/apply/main.jsp"
btime = "2018年3月1日-3月14日。"
Case 20
xm = "温州大学"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = 5 * cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.wzu.edu.cn/info/1010/1702.htm"
bm = "http://swyt.wzu.edu.cn:8088/apply/main.jsp"
btime = "2018年3月4日上午9:00至3月20日下午16:00；"
Case 21
xm = "浙江外国语学院"
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://swyt.zisu.edu.cn/newsShow.aspx?newsCate=9ZaCYPKsse9I1m4n_et1fQ==&newsID=59"
bm = "http://swyt.zisu.edu.cn/"
btime = "网上报名时间：2018年2月12日至3月9日；"
Case 22
xm = "宁波工程学院"
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
btime = "2018年3月1日—3月16日；"
Case 23
xm = "衢州学院"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://xsc.qzu.zj.cn/zs/info_1261.aspx"
bm = "http://xsc.qzu.zj.cn:8081/apply/main.jsp"
btime = "网上报名时间：2018年3月5日至3月18日；；"
Case 24
xm = "浙江水利水电学院"
cxk = ag * 10 + bg * 7 + cg * 4 + 1 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zjc.zjweu.edu.cn/zhaosheng/f6/2b/c1620a63019/page.htm"
bm = "http://swyt.zjweu.edu.cn/"
btime = "网上报名：2018年3月5日-3月25日；"
Case 25
xm = "丽水学院"
cxk = ag * 10 + bg * 7 + cg * 4 + 1 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.lsu.edu.cn/2018/0207/c619a253907/page.htm"
bm = "http://47.97.34.107/apply/main.jsp"
btime = "2018年3月4日（农历正月十七）—3月18日（农历二月初二）"
Case 26
xm = "温州肯恩大学"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk/1.5
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.wku.edu.cn/zh-hans/2018/02/20183in1admissions/"
bm = "http://121.40.33.186:810/"
btime = "网上报名时间：2018 年3月5日 9:00至2018年3月26日 17:00；"
Case 27
xm = "宁波诺丁汉大学"
cxk = 100 - 5 * (10 - ag)
xkf = cxk
xb = 0.1
zb = 0.3
examb = 0.6
wz = "https://www.nottingham.edu.cn/cn/study/undergraduate/policy-and-requirements/2018-direct-entry.aspx"
bm = "http://triunity.nottingham.edu.cn/apply/main.jsp"
btime = "网上报名时间：2018年3月15日上午9:00至5月9日下午16:00"
Case 28
xm = "浙江大学城市学院"
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
btime = "网上报名时间：2018年3月1日至3月11日；"
Case 29
xm = "浙江大学宁波理工学院"
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
btime = "2018年3月4日上午9:00至29日下午17:00;"
Case 30
xm = "浙江树人学院（浙江树人大学）"
cxk = ag * 10 + bg * 7 + cg * 3 + dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjsru.edu.cn/Cms/CmsDetail/182"
bm = "https://swyt.zjsru.edu.cn/apply/main.jsp"
btime = "2018年3月2日 - 3月19日。"
Case 31
xm = "浙江越秀外国语学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.yxc.cn/xsrx_1.aspx?newsid=619"
bm = "http://zs.yxc.cn/3/index.aspx"
btime = "2018年2月26日至3月18日（逾期不报）；"
Case 32
xm = "宁波大红鹰学院"
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
btime = "2018年2月23日—3月22日。"
Case 33
xm = "温州医科大学仁济学院"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaosheng.wmu.edu.cn/Art/Art_1/Art_1_3080.aspx"
bm = "http://swyt.rjxy.wmu.edu.cn/apply/main.jsp"
btime = "2018年2月26日上午8:30-3月18日下午16:00。"
Case 34
xm = "浙江中医药大学滨江学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://swytbm.zcmu.edu.cn/swytyunbjxy/apply/main.jsp"
bm = "http://swytbm.zcmu.edu.cn/swytyunbjxy/apply/main.jsp"
btime = "2018年3月4日9：00-2018年3月17日16：00止"
Case 35
xm = "中国计量大学现代科技学院"
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=2144"
bm = "http://swytxk.cjlu.edu.cn/apply/main.jsp"
btime = "2018年3月5日9:00至3月25日16:00；"
Case 36
xm = "杭州师范大学钱江学院"
cxk = ag * 10 + bg * 7 + cg * 5
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://qjzs.hznu.edu.cn/shownews.asp?id=816"
bm = "http://qjzsbm.hznu.edu.cn/zsbm_show/login.asp"
btime = "报名时间：2018年3月3日至3月25日"
Case 37
xm = "温州商学院"
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
btime = "2018年3月5日—3月31日"
Case 38
xm = "同济大学浙江学院"
cxk = ag * 10 + bg * 5 + cg * 3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://www.tjzj.edu.cn/zhaosheng.php?pid=745&cid=763&id=343313"
bm = "http://122.225.19.18:8081/apply/main.jsp"
btime = "2018年3月13日上午9：00至3月22日下午16：00。"
Case 39
xm = "上海财经大学浙江学院（主页无法访问）"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.shufe-zj.edu.cn/index.aspx?lanmuid=63&sublanmuid=65&id=6315"
bm = "http://swyt.shufe-zj.edu.cn/apply/main.jsp"
btime = "报名时间：2017年2月21日9:00—3月3日15:00"
Case 40
xm = "杭州医学院"
cxk = ag * 10 + bg * 7 +cg*3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.hmc.edu.cn/art/2018/2/8/art_2399_132743.html"
bm = ""
btime = "网上报名时间为2018年3月12日上午9:00～3月22日下午16:00"
Case 41
xm = "浙江工业大学之江学院"
cxk = ag * 10 + bg * 5 +cg*2
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zjc.zjut.edu.cn/archive/1059.html"
bm = "http://zs.zjc.zjut.edu.cn"
btime = "2018年3月5日上午9:00-3月19日下午16:00；"
Case 42
xm = "宁波大学科学技术学院 "
cxk = ag * 10 + bg * 7 +cg*3
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://zs.ndky.edu.cn/prospectuses/7427.jhtml"
bm = "http://swyt.ndky.edu.cn/apply/main.jsp"
btime = "2018年3月1日上午10:00至3月26日下午15:00；"
Case 43
xm = "杭州电子科技大学信息工程学院"
cxk = ag * 10 + bg * 5 +cg*2
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://www.hziee.edu.cn/index.php?c=Index&a=news_detail&catid=453&id=2812&web=bkszs"
bm = "http://swyt.hziee.edu.cn/"
btime = "2018年3月5日—2018年3月25日；"
Case 44
xm = "浙江财经大学东方学院"
cxk = ag * 10 + bg * 8 +cg*5+dg*2
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zufedfc.edu.cn/index.php?mod=show&mid=3&pid=193&id=2728"
bm = "http://zsbm.zufedfc.edu.cn:8080/apply/main.jsp"
btime = "2018年3月1日上午9：00至3月22日下午16：00"
Case 45
xm = "绍兴文理学院元培学院"
cxk = ag * 10 + bg * 8 +cg*7+dg*6
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsxx.ypc.edu.cn/info/1002/1424.htm"
bm = "http://zsxx.ypc.edu.cn/"
btime = "网上报名：2018年3月1日至3月25日；"
Case 46
xm = "温州大学瓯江学院"
cxk = ag * 10 + bg * 6 +cg*3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.ojc.zj.cn/Art/Art_383/Art_383_61403.aspx"
bm = "http://ojzs.cnvp.com.cn/apply/main.jsp"
btime = "2018年2月25日上午9:00至3月15日下午16:00。"
Case Else
MsgBox "ERROR：找不到此数据", vbExclamation, "系统消息"
End Select
End Sub

