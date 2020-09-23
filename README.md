<div align="center">

## Change MSN Messenger Name


</div>

### Description

This Code Lets you to change your name on msn messenger
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Justin Middelton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/justin-middelton.md)
**Level**          |Beginner
**User Rating**    |3.0 (15 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/justin-middelton-change-msn-messenger-name__1-33792/archive/master.zip)





### Source Code

```
'Make sure u add the Messenger Type Library by going to refreneces or if your on vb3,vb4 go to tools then refreneces on vb5,vb6 go to project then refrences and add messenger type library
'Then double click on the form and go to Declerations and put the following code in
Dim MsN As New MsgrObject
'Now put 1 text area on the form and 1 Command Button and put the following Code
Private Sub Command1_Click()
MsN.Services(0).FriendlyName = Text1.Text
MsN.LocalState = MSTATE_INVISIBLE
MsN.LocalState = MSTATE_ONLINE
End Sub
'Thank you Hope you enjoy:)
```

