Attribute VB_Name = "Module2"
Public k As Integer
Public tag As Integer
Public Sub sjk()
Select Case xh
Case 1
xm = "浙江工业大学"
If ag < 5 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = 100 - 5 * (10 - ag) + 2 * bg
xkf = cxk
End If
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zjut.edu.cn/zsnews/html/n1753.html"
bm = "http://zs.zjut.edu.cn/swytyun/apply/main.jsp"
btime = "网上报名时间为2017年2月20日上午9:00至3月2日下午16:00；邮寄材料接收从报名时起至3月2日截止（以当地邮戳为准）"
Case 2
xm = "浙江师范大学"
If ag < 3 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjnu.edu.cn/2017/0123/c101a118137/page.htm"
bm = "http://zsb.zjnu.edu.cn/apply/main.jsp"
btime = "报名时间：2017年2月22日至3月4日"
Case 3
xm = "宁波大学"
If ag < 3 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = 100 - 5 * (10 - ag)
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zsb.nbu.edu.cn/Article/Index/219"
bm = "http://zsb.nbu.edu.cn/Students/Login"
btime = "网上报名、材料上传时间：2017年2月25日上午10:00至3月10日下午15:00；"
Case 4
xm = "杭州电子科技大学"
cxk = ag * 15 + bg * 10 + cg * 5
xkf = cxk / 1.5
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zhaosheng.hdu.edu.cn/art.php?aid=1255"
bm = "http://swyt.hdu.edu.cn/"
btime = "网上报名：2017年2月8日至3月8日；"
Case 5
xm = "浙江工商大学"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaoban.zjsu.edu.cn/html/2017-1/201712390751.htm"
bm = "http://zhaoban.zjsu.edu.cn/swyt/index.asp"
btime = "网上报名时间：2017年2月19日至3月1日下午16:00；"
Case 6
xm = "浙江理工大学"
If ag * 15 + bg * 9 + cg * 3 < 100 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk
End If
examb = 0.5
wz = "http://zs.zstu.edu.cn/?p=read&aid=1759"
bm = "http://120.55.84.14/zjlg/Login.aspx"
btime = "报名时间为即日起至2017年3月1日。"
Case 7
xm = "温州医科大学"
If ag < 5 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = 100 - 5 * (10 - ag)
xkf = cxk
End If
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zhaosheng.wmu.edu.cn/Art/Art_1/Art_1_1329.aspx"
bm = "http://swyt.wmu.edu.cn/apply/main.jsp"
btime = "考生报名：2017年2月10日上午8:30-2月26日下午16:00；"
Case 8
xm = "浙江海洋大学"
If ag < 2 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjou.edu.cn/info/1047/3021.htm"
bm = "http://swyt.zjou.edu.cn/user.asp"
btime = "网上报名时间：2017年2月19日—2017年3月9日；"
Case 9
xm = "浙江农林大学"
If ag < 3 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zafu.edu.cn/info/1016/2644.htm"
bm = "http://swyt.zafu.edu.cn/index.php/Login/index.html"
btime = "网上报名时间：2017年2月21日至3月6日下午4：00。"
Case 10
xm = "浙江中医药大学"
If ag < 5 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
End If
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.zcmu.edu.cn/news_show.asp?id=927"
bm = "http://swytbm.zcmu.edu.cn/stu_login.aspx"
btime = "网上报名和书面材料邮寄时间：2017年3月1日9：00-2017年3月15日16：00止"
Case 11
xm = "中国计量大学"
If ag < 2 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * dg
xkf = cxk
End If
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=1763"
bm = "http://swytzs.cjlu.edu.cn:8080/SWYTJL/AppMain.jsp"
btime = "网上报名时间：2017年3月3日9:00至3月22日16:00；"
Case 12
xm = "浙江万里学院"
cxk = ag * 10 + bg * 6 + cg * 3 + 1 * dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zswnew.zwu.edu.cn/bkszsw/72/ca/c4804a94922/page.htm"
bm = "https://swyt.zwu.edu.cn/renderLogin.do;JSESSIONID=d08cd28b-bbfc-49b7-88b8-3128d07f9fa3"
btime = "网上报名时间：2017年3月1日—3月20日。"
Case 13
xm = "浙江科技学院(普通类）"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsb.zust.edu.cn/IndexPage!wzxq.htm?id=3AA16DD86E67476CAE5820141547B6D6"
bm = "http://zsb.zust.edu.cn/IndexPage!login.htm"
btime = "报名时间：2017年2月7日-3月6日。"
Case 14
xm = "浙江财经大学"
If (ag * 10 + bg * 5 + cg * 2) < 70 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = ag * 10 + bg * 5 + cg * 2
xkf = cxk
End If
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zufe.edu.cn/info/1002/1623.htm"
bm = "http://swyt.zufe.edu.cn/apply/main.jsp"
btime = "网上报名：2017年2月20日上午9:00至3月1日下午16:00；"
Case 15
xm = "嘉兴学院"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://admission.zjxu.edu.cn/news/7080/read.shtml"
bm = "http://210.33.29.129/"
btime = "2017年2月20日—3月15日"
Case 16
xm = "杭州师范大学"
cxk = ag * 10 + bg * 5 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://bkzs.hznu.edu.cn/Details/20170125/410924201701250924378326.html"
bm = "http://hzsf.vnet1000.com"
btime = "2017年2月1日-3月3日"
Case 17
xm = "湖州师范学院"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = cxk / 1.5
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.hutc.zj.cn/2014/Item/Show.asp?m=1&d=2688"
bm = "http://swyt.zjhu.edu.cn/apply/main.jsp"
btime = "2017年2月22日-3月5日。"
Case 18
xm = "绍兴文理学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.usx.edu.cn/info/1010/2541.htm"
bm = "https://zsbm.usx.edu.cn/w3/index.html"
btime = "2017年2月22日-3月5日。"
Case 19
xm = "台州学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.zsjy.tzc.edu.cn/articles/1378.html"
bm = "http://swyt.tzc.edu.cn/apply/main.jsp"
btime = "2017年2月16日-3月1日。"
Case 20
xm = "温州大学"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = 5 * cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.wzu.edu.cn/Art/Art_5/Art_5_531.aspx"
bm = "http://swyt.wzu.edu.cn:8088/apply/main.jsp"
btime = "2017年2月20日上午9:00至3月2日下午16:00；"
Case 21
xm = "浙江外国语学院"
cxk = ag * 10 + bg * 6 + cg * 2
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://zsjy.zisu.edu.cn/zhaosheng/article.asp?id=1058&?typeid=2"
bm = "http://swyt.zisu.edu.cn/"
btime = "网上报名时间：2017年2月10日至3月9日；"
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
wz = "http://zs.nbut.edu.cn/index.php?s=/Home/Article/detail/id/692.html"
bm = "http://zs.nbut.edu.cn/index.php?s=/Home/Index/login.html"
btime = "网上报名时间：2017年2月20日—3月3日；"
Case 23
xm = "衢州学院"
cxk = ag * 10 + bg * 7 + cg * 4
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://xsc.qzu.zj.cn/zs/info_1193.aspx"
bm = "http://xsc.qzu.zj.cn:8081/apply/main.jsp"
btime = "网上报名时间：2017年2月17日至3月2日；"
Case 24
xm = "浙江水利水电学院"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zjc.zjweu.edu.cn/zhaosheng/d3/35/c1620a54069/page.htm"
bm = "http://swyt.zjweu.edu.cn/"
btime = "网上报名：2017年2月22日-3月8日；"
Case 25
xm = "丽水学院"
cxk = ag * 10 + bg * 7 + cg * 4 + 1 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zsw.lsu.edu.cn/2017/0124/c590a210179/page.htm"
bm = "http://swyt.lsu.edu.cn/apply/main.jsp"
btime = "2017年2月12日（农历正月十六）—2月25日（正月二十九）"
Case 26
xm = "温州肯恩大学"
cxk = ag * 15 + bg * 9 + cg * 3
xkf = 5 * cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://www.wku.edu.cn/zh-hans/2017/01/wenzhoukenendaxue2017niansanweiyiti-zonghepingjiazhaoshengzhangcheng/"
bm = "http://121.40.33.186:810/"
btime = "网上报名时间：2017年3月6日9:00至2017年3月27日17:00；"
Case 27
xm = "宁波诺丁汉大学"
If ag < 7 Then
MsgBox "未符合该校最低报名条件"
Else
cxk = 100 - 5 * (10 - ag)
xkf = cxk
End If
xb = 0.1
zb = 0.3
examb = 0.6
wz = "http://www.nottingham.edu.cn/cn/study/undergraduate/policy-and-requirements/2017-direct-entry.aspx"
bm = "http://125.111.163.252:8080/apply/main.jsp"
btime = "2017年3月13日09：00-5月4日16：00"
Case 28
xm = "浙江大学城市学院"
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
btime = "网上报名时间：2017年2月24日上午9:00至3月6日下午16:00；"
Case 29
xm = "浙江大学宁波理工学院"
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
btime = "2017年2月24日上午9:00至3月12日下午17:00;"
Case 30
xm = "浙江树人学院（浙江树人大学）"
cxk = ag * 10 + bg * 7 + cg * 3 + dg
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.zjsru.edu.cn/Cms/CmsDetail/154?20160701ba0b5679"
bm = "http://swyt.zjsru.edu.cn/apply/main.jsp"
btime = "网上报名时间：2017年2月17日 - 3月7日。"
Case 31
xm = "浙江越秀外国语学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.zyufl.edu.cn/xsrx_1.aspx?newsid=557"
bm = "http://zs.yxc.cn/3/index.aspx"
btime = "网上报名时间：2017年2月16日至3月16日（逾期不报）；"
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
wz = "http://www.dhyedu.com/zsb/show.htm?id=11175"
bm = "http://zsb.nbdhyu.edu.cn/"
btime = "网上报名时间：2017年2月10日—3月20日。"
Case 33
xm = "温州医科大学仁济学院"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://rjxy.wmu.edu.cn/zhaosheng/view.jsp?id0=&id=z0iyato2yn"
bm = "http://swyt.rjxy.wmu.edu.cn/apply/main.jsp"
btime = "2017年2月10日-2月26日。"
Case 34
xm = "浙江中医药大学滨江学院"
cxk = ag * 10 + bg * 8 + cg * 6 + 4 * dg
xkf = cxk
xb = 0.15
zb = 0.3
examb = 0.55
wz = "http://zsb.zcmu.edu.cn/news_show.asp?id=928"
bm = "http://swytbm.zcmu.edu.cn/stu_login.aspx"
btime = "2017年3月1日9：00-2017年3月15日16：00止"
Case 35
xm = "中国计量大学现代科技学院"
cxk = ag * 10 + bg * 9 + cg * 8 + 7 * (10 - ag - bg - cg)
xkf = cxk
xb = 0.15
zb = 0.35
examb = 0.5
wz = "http://zs.cjlu.edu.cn/Detail?id=1762"
bm = "http://swytzs.cjlu.edu.cn:8080/SWYTXDKJ/AppMain.jsp"
btime = "网上报名时间：2017年3月3日9:00至3月22日16:00；"
Case 36
xm = "杭州师范大学钱江学院"
cxk = ag * 10 + bg * 7 + cg * 5
xkf = cxk
xb = 0.1
zb = 0.4
examb = 0.5
wz = "http://qjzs.hznu.edu.cn/shownews.asp?id=758"
bm = "http://qjzsbm.hznu.edu.cn/zsbm_show/login.asp"
btime = "报名时间：2017年2月11日至3月5日。"
Case 37
xm = "温州商学院"
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
btime = "网上报名：2017年2月27日—3月26日"
Case 38
xm = "同济大学浙江学院"
cxk = ag * 10 + bg * 5 + cg * 3
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://www.tjzj.edu.cn/zhaosheng.php?pid=745&cid=763&id=341212"
bm = "http://122.225.19.18:8081/apply/main.jsp"
btime = "2017年2月25日上午9：00至3月6日下午16：00。"
Case 39
xm = "上海财经大学浙江学院"
cxk = ag * 10 + bg * 5
xkf = cxk
xb = 0.2
zb = 0.3
examb = 0.5
wz = "http://zs.shufe-zj.edu.cn/index.aspx?lanmuid=63&sublanmuid=65&id=6315"
bm = "http://swyt.shufe-zj.edu.cn/apply/main.jsp"
btime = "报名时间：2017年2月21日9:00—3月3日15:00"
Case Else
MsgBox "ERROR：找不到此数据", vbExclamation, "系统消息"
End Select
End Sub

