value=RandomNumber (0,100)
'user="test"
newuser = "test" & value
'msgbox newuser

SystemUtil.CloseProcessByName("IEXPLORE.exe")
DataTable.ImportSheet "C:\Temp\PC_UFT_Test.dat",1, "Global"

url = DataTable.RawValue("URL", 1)
username = DataTable.RawValue("USERNAME",1)

SystemUtil.Run "iexplore.exe", url

wait 10
Browser("Advantage Shopping").Page("Advantage Shopping_9").Sync

if(Browser("Advantage Shopping").Page("Advantage Shopping_9").WebElement("DEMO").GetROProperty("innertext")="DEMO") then
	Reporter.ReportEvent micPass,"Launch Home screen","Launch Home Screen Successfully with correct title"
else
	Reporter.ReportEvent micFail,"Launch Home screen","Launch Home Screen fail with incorrect title"
	Browser("Advantage Shopping").CloseAllTabs
	ExitTest
End if

If (Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Layer_1").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Layer_1").Click
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Layer_1").Click
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("CREATE NEW ACCOUNT").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("CREATE NEW ACCOUNT").Click
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("CREATE NEW ACCOUNT").Click
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit").Set newuser
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit").Set username
End If

 If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_2").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_2").Set "tracy@devops.com"
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_2").Set "tracy@devops.com"
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_3").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_3").SetSecure "58d4216e1c120b53b1a71a84c51b9b5f"
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_3").SetSecure "58d4216e1c120b53b1a71a84c51b9b5f"
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_4").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_4").SetSecure "58d4217d89fb2d152e77d8c49157fe12"
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebEdit("WebEdit_4").SetSecure "58d4217d89fb2d152e77d8c49157fe12"
End If
 
 If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("I agree to the www.AdvantageOn").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("I agree to the www.AdvantageOn").Click
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("I agree to the www.AdvantageOn").Click
End If


If (Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("REGISTER").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("REGISTER").Click
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_2").WebElement("REGISTER").Click
	
End If
wait 3
If (Browser("Advantage Shopping").Page("Advantage Shopping_3").WebElement("SpeakersImg").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_3").WebElement("SpeakersImg").Click
else
    wait 8
	Browser("Advantage Shopping").Page("Advantage Shopping_3").WebElement("SpeakersImg").Click
End If


If (Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("BUY NOW").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("BUY NOW").Click
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("BUY NOW").Click
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_4").WebButton("ADD TO CART").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_4").WebButton("ADD TO CART").Click
else
    wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping_4").WebButton("ADD TO CART").Click
End If


wait 5


If (Browser("Advantage Shopping").Page("Advantage Shopping_10").WebElement("WebElement").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_10").WebElement("WebElement").Click
else
    wait 6
	Browser("Advantage Shopping").Page("Advantage Shopping_10").WebElement("WebElement").Click
End If

wait 2
If (Browser("Advantage Shopping").Page("Advantage Shopping_11").WebButton("CHECKOUT ($200.00)").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_11").WebButton("CHECKOUT ($200.00)").Click
else
    wait 6
	Browser("Advantage Shopping").Page("Advantage Shopping_11").WebButton("CHECKOUT ($200.00)").Click
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebButton("NEXT").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebButton("NEXT").Click
else
    wait 6
	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebButton("NEXT").Click
End If




'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").Image("Master credit").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").Image("Master credit").Click
'else
'    wait 6
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").Image("Master credit").Click
'End If
'
'
'
'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Card number").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Card number").Click
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Card number").Click
'End If
'
'
'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit").Set "123456789012"
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit").Set "123456789012"
'End If
'
'
'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("CVV number").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("CVV number").Click
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("CVV number").Click
'End If
'
'Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit_2").Set "2353"


'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Card number field is required").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Card number field is required").Click
 '   Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit").Set "123456789012"
'End If



'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("CVV number field is required").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("CVV number field is required").Click
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("CVV number field is required").Click
'End If
'
'Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit_2").Set "123"
'
'
'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Cardholder name").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Cardholder name").Click
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("Cardholder name").Click
'End If
'
'Browser("Advantage Shopping").Page("Advantage Shopping_5").WebEdit("WebEdit_3").Set "tracy"
'
'If (Browser("Advantage Shopping").Page("Advantage Shopping_8").WebList("select").Exist) Then
'    Browser("Advantage Shopping").Page("Advantage Shopping_8").WebList("select").Select "2023"
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_8").WebList("select").Select "2023"
'End If
'
'
'If (Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("PAY NOW").Exist) Then
'    
'    Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("PAY NOW").Click
'else
'    wait 4
'	Browser("Advantage Shopping").Page("Advantage Shopping_5").WebElement("PAY NOW").Click
'End If
'

If (Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("ORDER PAYMENT").Exist) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("ORDER PAYMENT").Click
else
    wait 6
    Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("ORDER PAYMENT").Click
End If
 @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping 12").WebElement("ORDER PAYMENT")_;_script infofile_;_ZIP::ssf147.xml_;_
 If (Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("SafePay username").Exist) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("SafePay username").Click
else
    wait 6
    Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("SafePay username").Click
End If
 @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping 12").WebElement("SafePay username")_;_script infofile_;_ZIP::ssf148.xml_;_
 
 If (Browser("Advantage Shopping").Page("Advantage Shopping_12").WebEdit("WebEdit").Exist) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_12").WebEdit("WebEdit").Set "tracy"
else
    wait 6
    Browser("Advantage Shopping").Page("Advantage Shopping_12").WebEdit("WebEdit").Set "tracy"
End If

 If (Browser("Advantage Shopping").Page("Advantage Shopping_12").WebEdit("WebEdit_2").Exist) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_12").WebEdit("WebEdit_2").SetSecure "5b83874a59414e99802d9bb5b9ce1928fe8ca12e0a49744bbe375168"
else
    wait 6
    Browser("Advantage Shopping").Page("Advantage Shopping_12").WebEdit("WebEdit_2").SetSecure "5b83874a59414e99802d9bb5b9ce1928fe8ca12e0a49744bbe375168"
End If

 If (Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("PAY NOW").Exist) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("PAY NOW").Click
else
    wait 6
    Browser("Advantage Shopping").Page("Advantage Shopping_12").WebElement("PAY NOW").Click
End If
 @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping 12").WebEdit("WebEdit")_;_script infofile_;_ZIP::ssf149.xml_;_
 @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping 12").WebEdit("WebEdit 2")_;_script infofile_;_ZIP::ssf150.xml_;_
 @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping 12").WebElement("PAY NOW")_;_script infofile_;_ZIP::ssf151.xml_;_
 @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping 12").WebElement("PAY NOW")_;_script infofile_;_ZIP::ssf154.xml_;_
wait 6
If (Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("WebElement").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("WebElement").Click
else
    wait 6
	Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("WebElement").Click
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("My Orders").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("My Orders").Click
else
    wait 6
	Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("My Orders").Click
End If


If (Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("Layer_1").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("Layer_1").Click
else
    wait 3
	Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("Layer_1").Click
End If

If (Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("Sign out").Exist) Then
    Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("Sign out").Click
else
    wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping_6").WebElement("Sign out").Click
End If


wait 3
Browser("Advantage Shopping").CloseAllTabs

