' Context
AIUtil.SetContext SAPGuiSession("micclass:=SAPGuiSession")

' Search T-Code
AIUtil("combobox", "More").Type "/nva03" + vbCrLf
AIUtil("text_box", "Order:").Type "4040"
AIUtil("button", "Search").Click

' Text Validation
AIUtil.FindText("Delivery Block:").CheckExists True

' Navigate to Item detail
AIUtil.FindText("Reason for rejection").Click

' Grab value
Parameter ("searchValue") = AIUtil.FindTextBlock(micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock("Cust. Reference:")).GetValue

'MsgBox ("searchValue: " & searchValue)

' Exit
AIUtil.FindText("Exit").Click
AIUtil.FindText("Exit").Click
AIUtil("button", "Yes").Click

