<%
const DEF_UserEnableUserTitle = 1
const DEF_UserUserTitleNeedLevel = 8
const LMT_UserNameEnableEnglishWords = 1
const LMT_UserNameEnableChineseChar = 1
const LMT_UserNameEnableChineseWords = 1
const DEF_User_RegPoints = 0
const LMT_EnableRegNewUsers = 1
const DEF_ShortestUserName = 2
const DEF_RegNewUserTotalRestTime = 10
const DEF_UserNewRegAttestMode = 0
const DEF_UserActivationExpiresDay = 30
const DEF_User_GetPassMode = 2
Dim DEF_UserLevelString,DEF_UserLevelNum,DEF_UserLevelPoints
DEF_UserLevelString = Array("新手上路","论坛游民","论坛游侠","职业侠客","大侠","骑士","圣骑士","精灵","精灵王","风云使者","光明使者","天使","大天使","精灵使","法师","大法师","法王","老法王","天神","天王","法老")
DEF_UserLevelPoints = Array(0,12,25,50,80,150,250,400,700,1000,1500,2500,5000,8000,12000,20000,30000,40000,50000,60000,99999)
DEF_UserLevelNum = Ubound(DEF_UserLevelString,1)
Dim DEF_UserOfficerString,DEF_UserOfficerNum
DEF_UserOfficerString = Array("游侠","<span class=""greenfont"">退休版主</span>","<span class=""greenfont"">保留</span>","<span class=""greenfont"">保留</span>","<span class=""greenfont"">保留</span>","<span class=""greenfont"">保留</span>")
DEF_UserOfficerNum = Ubound(DEF_UserOfficerString,1)
const DEF_FiltrateUserNameString = ""
const DEF_UserShortestPassword = 5
const DEF_UserShortestPasswordMaster = 8
const Def_UserTestNumber = 1
const DEF_seller_email = ""
const DEF_seller_minpoints = 1
const DEF_seller_exchangescale = 1
%>
