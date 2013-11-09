Attribute VB_Name = "modETC"
'---------------------------------------------------------------------------------------
' Module    : modETC
' DateTime  : 2012-08-05 20:06
' Author    : PC1
' Purpose   :
'---------------------------------------------------------------------------------------
'프로젝트 설명
'이 버전은 베타 버전으로, 실행시 로그 파일이 생성됩니다.
'이것이 싫으신 분은 모듈 modMain의 상수 부분에서 DEBUG_VERSION 부분의 값을 False로 바꿔 주시면 됩니다.
'
'참고사항
'로그 파일 저장방식을 RichTextBox 컨트롤을 이용하던 방식에서 직접 열어서 작업하는 방식으로 바꾸었습니다.
'또한 텍스트 파일 여는 방법 또한 개선 중입니다.(직접 여는 방식으로)
'!경고!
'이 프로그램의 소스 코드는 허락 없이 도용할 수 없습니다!
'
'Copyright YJSoft(yyj9411@naver.com). All rights Reserved.
'
'아래 모듈은 각각 모듈 위에 저작권 정보가 표기되어 있습니다.
'

'디스크 공간 구하는 함수 By HappyBono(http://www.happybono.net/285)
'CC BY-NC-ND
'저작자표시-비영리-변경금지
'http://creativecommons.org/licenses/by-nc-nd/2.0/kr/ 참고

Public Function GetDiskFreeSpace(strDeviceID As String) As String
Dim oWMI As Object
Dim oLDK As Object

Set oWMI = GetObject("winmgmts:")

Set oLDK = oWMI.Get("Win32_LogicalDisk.DeviceID=" _
& Chr(39) & strDeviceID & Chr(58) & Chr(39))

GetDiskFreeSpace = objLogicalDisk.FreeSpace & " byte"

Set oWMI = Nothing
Set oLDK = Nothing

End Function
