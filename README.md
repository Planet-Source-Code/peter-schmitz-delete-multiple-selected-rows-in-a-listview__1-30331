<div align="center">

## Delete multiple selected rows in a listview


</div>

### Description

Always wanted to delete multiple rows in a listview by using checkboxes to select which ones to delete? Then this is the simple solution.

Might work on earlier versions of VB, too.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter Schmitz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-schmitz.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-schmitz-delete-multiple-selected-rows-in-a-listview__1-30331/archive/master.zip)





### Source Code

```
Private Sub cmdDelete_Click()
Dim i As Integer
 With Listview1
 ' The trick: We step backwards through
 ' the array.
 ' The reason you always get an 'out of
 ' bound' error is because at a certain
 ' point the value of i will equal 0,
 ' or be greater than the number of rows
 ' left. (We set i with the initial
 ' row.count, and then start deleting
 ' from that count).
 ' We avoid that by stepping
 ' backwards :)
 For i = .ListItems.Count To 1 Step -1
 If .ListItems(i).Checked Then
  .ListItems.Remove (i)
 End If
 Next i
 End With
End Sub
```

