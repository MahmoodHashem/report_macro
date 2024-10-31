' Module-level variable
Dim weekCounter As Long

Sub InitializeWeekCounter()
    ' Check if the cell has a value and initialize weekCounter
    Set ws = ThisWorkbook.Sheets("Number of Days")
    
   Dim password As String
    password = "12341"
    
   ws.Unprotect password:=password

   weekCounter = ws.Range("B1").Value

   ws.Protect password:=password
 
    
   
End Sub

Sub IncrementWeekCounter()
    weekCounter = weekCounter + 1
    ' ThisWorkbook.Sheets("Number of Days").Range("B1").Value = weekCounter ' Store back in the cell
    
     Set ws = ThisWorkbook.Sheets("Number of Days")
    
   Dim password As String
    password = "12341"
    
   ws.Unprotect password:=password

    ws.Range("B1").Value = weekCounter

   ws.Protect password:=password
End Sub

Sub AddDailyEarningsTable()

' اول از کاربر تاييدي مي گيريم که امروز جدول ديگري نساخته باشد

Dim response As VbMsgBoxResult
response = MsgBox("آيا اين جدول اولي براي امروز است؟(ساختن بيش از يک جدول در روز سيستم را برهم ميزند)", vbYesNo + vbQuestion, "Confirmation")

' اگر جواب بلي بود پروسه را انجام مي دهيم

If response = vbYes Then
InitializeWeekCounter
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("گزارش روزانه") ' Change to your daily sheet name
   
    ' Find the last used row in column A
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Set the starting point for the new table
    Dim startRow As Long
    startRow = lastRow + 4 ' Add a gap of one row between tables

    ' Define the headers in English
    Dim headers As Variant
    headers = Array("بخش", "تعداد مريض", "مجموع", "قابله", "داکتر 1", _
                    "داکتر 2", "داکتر 3", "% شفاخانه", "مصارف", _
                    "مبلغ", "دواخانه", "خريد دوا", "باقي دواخانه", _
                    "مفاد خالص")

    ' Set headers in the worksheet
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(startRow, i + 1).Value = headers(i)
    Next i

    ' Define the rows under "Ward" in English
    Dim wards As Variant
    wards = Array("بخش عاجل", "تکس اطاق", "فيس داخله", "ولادت", _
                  "فيس نسايي", "بخش دندان", "سنوگرافي", "لابراتوار", "تکس عمليات", " ", "مجموع")

    ' Set wards in the worksheet
    Dim j As Long
    For j = LBound(wards) To UBound(wards)
        ws.Cells(startRow + 1 + j, 1).Value = wards(j)
    Next j

    ' Define the range for the table
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 12, UBound(headers) + 1))
    
    
    ' Add borders to the table range
    With tblRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0 ' Automatic color
    End With

    ' Auto-fit columns for better visibility
    ws.Columns("A:N").AutoFit
    
     ' Store the value from the 8th column and 4th row in a variable
       ws.Cells(startRow + 11, 8).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
       
      ' اين فرمول ها براي جمع کردن فيصدي شفاخانه است يعني در ستون شفاخانه هر سطر سمت راست خود را جمع مي کند
      ws.Cells(startRow + 1, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 2, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 3, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 4, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 5, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 6, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 7, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 8, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ws.Cells(startRow + 9, 8).FormulaR1C1 = "=RC[-5]-SUM(RC[-4]:RC[-1])"
      ' /////////////////////////////////////////////////////////////
      
      ' اين ستون ها کل جمع مصارف، خريد دوا و دواخانه است که ستون مرتبط خود را جمع مي کند
      ws.Cells(startRow + 11, 10).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)" ' 4th row is startRow + 3, 8th column is 8
      ws.Cells(startRow + 11, 11).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ws.Cells(startRow + 11, 12).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ' //////////////////////////////////////////////////////////////
      
      ws.Cells(startRow + 11, 13).FormulaR1C1 = "=RC[-2]-RC[-1]"
      ws.Cells(startRow + 11, 14).FormulaR1C1 = "=RC[-6]-RC[-4]"
      
      ' اين فرمول ها براي سطر مجموع پايين است که ستون هاي مرتبط خود را جمع مي کند
      ws.Cells(startRow + 11, 2).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ws.Cells(startRow + 11, 3).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ws.Cells(startRow + 11, 4).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ws.Cells(startRow + 11, 5).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ws.Cells(startRow + 11, 6).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
      ws.Cells(startRow + 11, 7).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
    ' ////////////////////////////////////////////////////
      
        Dim lastTableRow As Long
    lastTableRow = startRow + UBound(wards) + 2 ' Last row after wards

    With ws.Range(ws.Cells(startRow - 1, 1), ws.Cells(startRow - 1, UBound(headers) + 1)) ' Adjust row as needed
        .Merge
        .Value = Slash(Shamsi()) ' Set the header value with the date
        .HorizontalAlignment = xlCenter ' Center align the text
        .ReadingOrder = xlRTL
        .Font.Bold = True ' Make the text bold
        .Font.Size = 14 ' Optional: Set font size
        .Interior.Color = RGB(0, 176, 240) ' Optional: Set background color
    End With
    With ws.Range(ws.Cells(startRow - 2, 1), ws.Cells(startRow - 2, UBound(headers) + 1)) ' Adjust row as needed
        .Merge
        .Value = DayWeek(Shamsi()) ' Set the header value with the date
        .ReadingOrder = xlRTL
        .HorizontalAlignment = xlCenter ' Center align the text
        .Font.Bold = True ' Make the text bold
        .Font.Size = 14 ' Optional: Set font size
        .Interior.Color = RGB(0, 176, 240) ' Optional: Set background color
    End With

    With ws.Range(ws.Cells(lastTableRow, 1), ws.Cells(lastTableRow, 12)) ' Merging first 10 columns
        .Merge
        .Value = "" ' Optional: Set a value for the merged cell
        .HorizontalAlignment = xlCenter ' Center align the text
        .Font.Bold = True ' Make the text bold
    End With
    
    With ws.Range(ws.Cells(lastTableRow, 13), ws.Cells(lastTableRow, 14))
        .Merge
        .FormulaR1C1 = "=R[-1]C[1]+R[-1]C"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
     End With
 IncrementWeekCounter
 Else
    MsgBox "درست است، شما قبلا يک جدول براي امروز ساخته ايد"
 End If
 
  
End Sub

Sub reportTable(ByVal empatient As Long, ByVal emTotal, ByVal emHos, ByVal bedPatient As Long, ByVal bedTotal As Long, ByVal bedHos As Long, ByVal inPatient As Long, ByVal inTotal As Long, ByVal inRah As Long, ByVal inShahim As Long, ByVal inWali As Long, ByVal inHos As Long, ByVal childPs As Long, ByVal childTotal As Long, ByVal childYek As Long, ByVal childHos As Long, ByVal womanPatients As Long, ByVal womanTotal As Long, ByVal womanYek As Long, ByVal womanHos As Long, ByVal dentistPateints As Long, ByVal dentistTotal As Long, ByVal dentistHos As Long, ByVal ultraPateints As String, ByVal ultraTotal As String, ByVal ultraYek As String, ByVal ultraRah As String, ByVal ultraShahim As String, ByVal ultraWali As String, ByVal ultraHos As String, ByVal labPateints As String, ByVal labTotal As String, ByVal labYek As String, ByVal labRah As String, ByVal labShahim As String, ByVal labWali As String, ByVal labHos As String, _
ByVal surP As Long, ByVal surTotal As Long, ByVal surHos As Long, ByVal exp As Long, ByVal phB As Long, ByVal phDB As Long, _
ByVal startDate As String, ByVal endDate As String)

Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("گزارش هفته وار") ' Change to your daily sheet name
   
   

    ' Find the last used row in column A
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Set the starting point for the new table
    Dim startRow As Long
    startRow = lastRow + 4 ' Add a gap of one row between tables

    ' Define the headers in English
    Dim headers As Variant
    headers = Array("بخش", "تعداد مريض", "مجموع", "قابله", "داکتر 1", _
                    "داکتر 2", "داکتر 3", "% شفاخانه", _
                    "مجموع مصارف هفته", "دواخانه", "خريد دوا", "باقي دواخانه", _
                    "مفاد خالص")

    ' Set headers in the worksheet
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(startRow, i + 1).Value = headers(i)
    Next i

    ' Define the rows under "Ward" in English
    Dim wards As Variant
    wards = Array("بخش عاجل", "تکس اطاق", "فيس داخله", "ولادت", _
                  "فيس نسايي", "بخش دندان", "سنوگرافي", "لابراتوار", "تکس عمليات", " ", "مجموع")

    ' Set wards in the worksheet
    Dim j As Long
    For j = LBound(wards) To UBound(wards)
        ws.Cells(startRow + 1 + j, 1).Value = wards(j)
    Next j

    ' Define the range for the table
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 12, UBound(headers) + 1))
    

    ' Add borders to the table range
    
    
  With ws.Range(ws.Cells(startRow - 1, 1), ws.Cells(startRow - 1, UBound(headers) + 1)) ' Adjust row as needed
        .Merge
        .Value = " گزارش هفته وار از تاريخ" & startDate & " الي " & endDate   ' Set the header value with the date
        .HorizontalAlignment = xlCenter ' Center align the text
        .ReadingOrder = xlRTL
        .Font.Bold = True ' Make the text bold
        .Font.Size = 14 ' Optional: Set font size
        .Interior.Color = RGB(0, 176, 240) ' Optional: Set background color
    End With
    
    
    With tblRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0 ' Automatic color
    End With
    
    
    With tblRange
     .HorizontalAlignment = xlCenter
    End With
    

    ' Auto-fit columns for better visibility
    ws.Columns("A:N").AutoFit
    
    ws.Cells(startRow + 11, 8).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)" ' formula for summing the hospital percentage earning
    ws.Cells(startRow + 11, 12).FormulaR1C1 = "=RC[-2]-RC[-1]"  ' formula for geting pure earning of phramacy
    ws.Cells(startRow + 11, 13).FormulaR1C1 = "=RC[-5]-RC[-4]"  ' formula for getting pure earning of hospital
    
    
    
 ' cells which are belong to emergency ward to thier according doers
 ws.Cells(startRow + 1, 2).Value = empatient
 ws.Cells(startRow + 1, 3).Value = emTotal
 ws.Cells(startRow + 1, 8).Value = emHos
 
 ' Cells which are belong to Bed taxes ward
    ws.Cells(startRow + 2, 2).Value = bedPatient
    ws.Cells(startRow + 2, 3).Value = bedTotal
    ws.Cells(startRow + 2, 8).Value = bedHos

' Celss which are belong to Interior ward under each of their doers
 ws.Cells(startRow + 3, 2).Value = inPatient
 ws.Cells(startRow + 3, 3).Value = inTotal
 ws.Cells(startRow + 3, 5).Value = inRah
 ws.Cells(startRow + 3, 6).Value = inShahim
 ws.Cells(startRow + 3, 7).Value = inWali
 ws.Cells(startRow + 3, 8).Value = inHos
 
 ws.Cells(startRow + 4, 2).Value = childPs
 ws.Cells(startRow + 4, 3).Value = childTotal
 ws.Cells(startRow + 4, 4).Value = childYek
 ws.Cells(startRow + 4, 8).Value = childHos
 
 ws.Cells(startRow + 5, 2).Value = womanPatients
  
ws.Cells(startRow + 5, 3).Value = womanTotal
ws.Cells(startRow + 5, 4).Value = womanYek
ws.Cells(startRow + 5, 8).Value = womanHos

ws.Cells(startRow + 6, 2).Value = dentistPateints
ws.Cells(startRow + 6, 3).Value = dentistTotal
ws.Cells(startRow + 6, 8).Value = dentistHos

ws.Cells(startRow + 7, 2).Value = ultraPateints
ws.Cells(startRow + 7, 3).Value = ultraTotal
ws.Cells(startRow + 7, 4).Value = ultraYek
ws.Cells(startRow + 7, 5).Value = ultraRah
ws.Cells(startRow + 7, 6).Value = ultraShahim
ws.Cells(startRow + 7, 7).Value = ultraWali
ws.Cells(startRow + 7, 8).Value = ultraHos

ws.Cells(startRow + 8, 2).Value = labPateints
ws.Cells(startRow + 8, 3).Value = labTotal
ws.Cells(startRow + 8, 4).Value = labYek
ws.Cells(startRow + 8, 5).Value = labRah
ws.Cells(startRow + 8, 6).Value = labShahim
ws.Cells(startRow + 8, 7).Value = labWali
ws.Cells(startRow + 8, 8).Value = labHos

ws.Cells(startRow + 9, 2).Value = surP
ws.Cells(startRow + 9, 3).Value = surTotal
ws.Cells(startRow + 9, 8).Value = surHos

ws.Cells(startRow + 11, 9).Value = exp
ws.Cells(startRow + 11, 10).Value = phB
ws.Cells(startRow + 11, 11).Value = phDB

 
 
End Sub

Sub CreateWeeklyReport()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalEarnings As Double
    Dim totalExpenses As Double
    Dim startRow As Long
    Dim daysCount As Long

    ' Set the daily report worksheet
    Set ws = ThisWorkbook.Sheets("گزارش روزانه")

  Set reportWs = ThisWorkbook.Sheets("گزارش هفته وار")


    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
     

    Dim satData(0 To 8, 0 To 6) As Variant
    Dim sunData(0 To 8, 0 To 6) As Variant
    Dim monData(0 To 8, 0 To 6) As Variant
    Dim tueData(0 To 8, 0 To 6) As Variant
    Dim wedData(0 To 8, 0 To 6) As Variant
    Dim thuData(0 To 8, 0 To 6) As Variant
    Dim friData(0 To 8, 0 To 6) As Variant
   
   
   'اين متغير ها بخاطر پاس دادن مصارف، دخل دواخانه و خريد دواي هفته  استفاده مي شود
    Dim exps As String
    Dim pharB As String
    Dim drugBuys As String
    
    exps = 0
    pharB = 0
    drugBuys = 0
    

    
    '  And weekCounter Mod 7 = 0

    If lastRow > 100  Then
    
    ' هر کدام اينها ميرود به خانه مورد نظر همان روز و کل همان سلول را داخل يک متغير ذخيره مي کند
        Set friPateince = ws.Cells(lastRow - 10, 2)
        Set thuPateince = ws.Cells(lastRow - 25, 2)
        Set wedPateince = ws.Cells(lastRow - 40, 2)
        Set tuePateince = ws.Cells(lastRow - 55, 2)
        Set monPateince = ws.Cells(lastRow - 70, 2)
        Set sunPateince = ws.Cells(lastRow - 85, 2)
        Set satPateince = ws.Cells(lastRow - 100, 2)
        Set satExpenses = ws.Cells(lastRow - 90, 10)
        Set satPharmacy = ws.Cells(lastRow - 90, 11)
        Set satDrugBuy = ws.Cells(lastRow - 90, 12)
        
        ' تاريخ روز جمعه يعني اول هفته را مي گيريم تا از تاريخ پنج شنبه کم کنيم و فاصله اين هفته را بيابيم
        Set startDate = ws.Cells(lastRow - 102, 1)
        Set endDate = ws.Cells(lastRow - 12, 1)
        
        

        Set friTotal = ws.Cells(lastRow - 10, 3)
        Set thuTotal = ws.Cells(lastRow - 25, 3)
        Set wedTotal = ws.Cells(lastRow - 40, 3)
        Set tueTotal = ws.Cells(lastRow - 55, 3)
        Set monTotal = ws.Cells(lastRow - 70, 3)
        Set sunTotal = ws.Cells(lastRow - 85, 3)
        Set sunExpenses = ws.Cells(lastRow - 75, 10)
        Set sunPharmacy = ws.Cells(lastRow - 75, 11)
        Set sunDrugBuy = ws.Cells(lastRow - 75, 12)
        Set satTotal = ws.Cells(lastRow - 100, 3)
        


        Set friYek = ws.Cells(lastRow - 10, 4)
        Set thuYek = ws.Cells(lastRow - 25, 4)
        Set wedYek = ws.Cells(lastRow - 40, 4)
        Set tueYek = ws.Cells(lastRow - 55, 4)
        Set monYek = ws.Cells(lastRow - 70, 4)
        Set monExpenses = ws.Cells(lastRow - 60, 10)
         Set monPharmacy = ws.Cells(lastRow - 60, 11)
         Set monDrugBuy = ws.Cells(lastRow - 60, 12)
        Set sunYek = ws.Cells(lastRow - 85, 4)
        Set satYek = ws.Cells(lastRow - 100, 4)
       


       Set friRah = ws.Cells(lastRow - 10, 5)
        Set thuRah = ws.Cells(lastRow - 25, 5)
        Set wedRah = ws.Cells(lastRow - 40, 5)
        Set tueRah = ws.Cells(lastRow - 55, 5)
        Set tueExpenses = ws.Cells(lastRow - 45, 10)
         Set tuePharmacy = ws.Cells(lastRow - 45, 11)
         Set tueDrugBuy = ws.Cells(lastRow - 45, 12)
        Set monRah = ws.Cells(lastRow - 70, 5)
        Set sunRah = ws.Cells(lastRow - 85, 5)
       Set satRah = ws.Cells(lastRow - 100, 5)

       Set friShahim = ws.Cells(lastRow - 10, 6)
        Set thuShahim = ws.Cells(lastRow - 25, 6)
        Set wedShahim = ws.Cells(lastRow - 40, 6)
        Set wedExpenses = ws.Cells(lastRow - 30, 10)
         Set wedPharmacy = ws.Cells(lastRow - 30, 11)
         Set wedDrugBuy = ws.Cells(lastRow - 30, 12)
        Set tueShahim = ws.Cells(lastRow - 55, 6)
        Set monShahim = ws.Cells(lastRow - 70, 6)
        Set sunShahim = ws.Cells(lastRow - 85, 6)
       Set satShahim = ws.Cells(lastRow - 100, 6)

        Set friWali = ws.Cells(lastRow - 10, 7)
        Set thuWali = ws.Cells(lastRow - 25, 7)
        Set thuExpenses = ws.Cells(lastRow - 15, 10)
         Set thuPharmacy = ws.Cells(lastRow - 15, 11)
         Set thuDrugBuy = ws.Cells(lastRow - 15, 12)
        Set wedWali = ws.Cells(lastRow - 40, 7)
        Set tueWali = ws.Cells(lastRow - 55, 7)
        Set monWali = ws.Cells(lastRow - 70, 7)
        Set sunWali = ws.Cells(lastRow - 85, 7)
       Set satWali = ws.Cells(lastRow - 100, 7)

       Set friHos = ws.Cells(lastRow - 10, 8)
       Set friExpenses = ws.Cells(lastRow, 10)
        Set friPharmacy = ws.Cells(lastRow, 11)
        Set friDrugBuy = ws.Cells(lastRow, 12)
        Set thuHos = ws.Cells(lastRow - 25, 8)
        Set wedHos = ws.Cells(lastRow - 40, 8)
        Set tueHos = ws.Cells(lastRow - 55, 8)
        Set monHos = ws.Cells(lastRow - 70, 8)
        Set sunHos = ws.Cells(lastRow - 85, 8)
       Set satHos = ws.Cells(lastRow - 100, 8)
       
       
      ' حاصل جمعه مصارف، دخل دواخانه و باقي دواخانه تمام هفته
      exps = satExpenses.Value + sunExpenses.Value + monExpenses.Value + tueExpenses.Value + wedExpenses.Value + thuExpenses.Value + friExpenses.Value
      pharB = satPharmacy.Value + sunPharmacy.Value + monPharmacy.Value + tuePharmacy.Value + wedPharmacy.Value + thuPharmacy.Value + friPharmacy.Value
      drugBuys = satDrugBuy.Value + sunDrugBuy.Value + monDrugBuy.Value + tueDrugBuy.Value + wedDrugBuy.Value + thuDrugBuy.Value + friDrugBuy.Value

       
    
       
    
     ' اين فور لوپ ها بخاطر ذخيره کردن داده هاي هر روز استفاده مي شود
     '  يعني مثلا ميرود به روز شنبه و از بخش عاجل شروع مي کند به سمت پايين تا تکس عمليات را داخل يک واريبل ذخيره مي کند
     ' بعد همين قسم ميرود به سمت راست تعداد مريض، مجموع، سهم هر داکتر، و فيصد شفاخانه را داخل يک واريبل ذخيره مي کند

     For i = 0 To 8
         satPateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         satTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         satYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         satRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         satShahim.Offset(i, 0).Interior.Color = RGB(107, 21, 230)
         satWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         satHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
         
         For j = 0 To 6
            satData(i, j) = satPateince.Offset(i, j) ' Example value assignment
         Next j
         
         
         

     Next i
     For i = 0 To 8
         sunPateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         sunTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         sunYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         sunRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         sunShahim.Offset(i, 0).Interior.Color = RGB(107, 21, 230)
         sunWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         sunHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
         
        For j = 0 To 6
            sunData(i, j) = sunPateince.Offset(i, j) ' Example value assignment
        Next j
        
     Next i
      For i = 0 To 8
         monPateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         monTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         monYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         monRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         monShahim.Offset(i, 0).Interior.Color = RGB(107, 21, 230)
         monWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         monHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
         
       For j = 0 To 6
            monData(i, j) = monPateince.Offset(i, j) ' Example value assignment
       Next j
        
         
      Next i
     
     For i = 0 To 8
        tuePateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         tueTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         tueYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         tueShahim.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         tueRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         tueWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         tueHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
         
         For j = 0 To 6
            tueData(i, j) = tuePateince.Offset(i, j) ' Example value assignment
        Next j
        
         
     Next i
     
     For i = 0 To 8
       wedPateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         wedTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         wedYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         wedRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         wedShahim.Offset(i, 0).Interior.Color = RGB(107, 21, 230)
         wedWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         wedHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
         
       For j = 0 To 6
            wedData(i, j) = wedPateince.Offset(i, j) ' Example value assignment
        Next j
        
         
        Next i
     
      For i = 0 To 8
       thuPateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         thuTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         thuYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         thuRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         thuShahim.Offset(i, 0).Interior.Color = RGB(107, 21, 230)
         thuWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         thuHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
         
         For j = 0 To 6
            thuData(i, j) = thuPateince.Offset(i, j) ' Example value assignment
        Next j
         
         
     Next i
       For i = 0 To 8
         friPateince.Offset(i, 0).Interior.Color = RGB(173, 216, 230) ' Store the value in the array
         friTotal.Offset(i, 0).Interior.Color = RGB(173, 21, 230)
         friYek.Offset(i, 0).Interior.Color = RGB(17, 21, 230)
         friRah.Offset(i, 0).Interior.Color = RGB(171, 121, 230)
         friShahim.Offset(i, 0).Interior.Color = RGB(107, 21, 230)
         friWali.Offset(i, 0).Interior.Color = RGB(17, 221, 230)
         friHos.Offset(i, 0).Interior.Color = RGB(171, 210, 230)
    
         For j = 0 To 6
            friData(i, j) = friPateince.Offset(i, j) ' Example value assignment
        Next j
         
     Next i
     
     
    Else
    MsgBox "براي تهيه گزارش هفته وار بايد هفته کامل باشد( از جمعه تا پنجشنبه)", vbCritical, "هفته تکميل نيست"
    
    End If
    
    ' اين آرايه براي ذخيره کردن داده هاي هر روز استفاده مي شود داده ها تمام گذارش بخش از عاجل شروع تا تکس عملسات
    ' تا تکس عمليات را با تمام سهم گيرنده هاي مربوطه آنها ذخيره مي کند
    Dim weeklyData(0 To 6) As Variant
    weeklyData(0) = satData
    weeklyData(1) = sunData
    weeklyData(2) = monData
    weeklyData(3) = tueData
    weeklyData(4) = wedData
    weeklyData(5) = thuData
    weeklyData(6) = friData
    
    
    
    ' اين متغير ها بخاطر پاس داده به گزارش هفته وار استفاده ميشود تا اينکه آنجا بتوانيم نمايش بدهيم
    
    Dim empateint As Long, emTotal As Long, emHos As Long
    Dim bedPatient As Long, bedTotal As Long, bedHos As Long
    Dim inPatient As Long, inTotal As Long, inRah As Long, inShahim As Long, inWali As Long, inHos As Long
    Dim childbirthPatients As Long, childbirthTotal As Long, childbirthHos As Long
    Dim womanPatients As Long, womanTotal As Long, womanYek As Long, womanHos As Long
    Dim dentistPateints As Long, dentistTotal As Long, dentistHos As Long
    Dim ultraPateints As Long, ultraTotal As Long, ultraYek As Long, ultraRah As Long, ultraShahim As Long, ultraWali As Long, ultraHos As Long
    Dim labPateints As Long, labTotal As Long, labYek As Long, labRah As Long, labShahim As Long, labWali As Long, labHos As Long
    Dim surgPatients As Long, surgTotal As Long, surgHos As Long
    Dim expenses As Long
    Dim pharmacyExpenses As Long, pharmacyEarnings As Long



    ' اين لوپ تمام گزارش هفته را از بخش عاجل با هم جمع ميکند تا معلوم شود اين هفته ""بخش عاجل"" چند مريض داشته و چندي سود کرده
        For i = 0 To 6
            empateint = empateint + weeklyData(i)(0, 0)
            emTotal = emTotal + weeklyData(i)(0, 1)
            emHos = emHos + weeklyData(i)(0, 6)
        Next i
        
        
     ' اين لوپ براي معلوم کردن تعداد مريض وسود همين هفته بخش ""تکس اطاق"" استفاده مي شود
        For i = 0 To 6
            bedPatient = bedPatient + weeklyData(i)(1, 0)
            bedTotal = bedTotal + weeklyData(i)(1, 1)
            bedHos = bedHos + weeklyData(i)(1, 6)
        Next i
        
      ' اين لوپ بخاطر جمع کردن تعداد مريض، سود کل، فيصدي هر داکتر و فيصدي شفاخانه بخش ""داخله"" را حساب ميکند
      For i = 0 To 6
        inPatient = inPatient + weeklyData(i)(2, 0)
        inTotal = inTotal + weeklyData(i)(2, 1)
        inRah = inRah + weeklyData(i)(2, 3)
        inShahim = inShahim + weeklyData(i)(2, 4)
        inWali = inWali + weeklyData(i)(2, 5)
        inHos = inHos + weeklyData(i)(2, 6)
    Next i

    ' Process childbirth data
    For i = 0 To 6
        childbirthPatients = childbirthPatients + weeklyData(i)(3, 0)
        childbirthTotal = childbirthTotal + weeklyData(i)(3, 1)
        childbirthHos = childbirthHos + weeklyData(i)(3, 6)
    Next i

    ' Process woman data
    For i = 0 To 6
        womanPatients = womanPatients + weeklyData(i)(4, 0)
        womanTotal = womanTotal + weeklyData(i)(4, 1)
        womanYek = womanYek + weeklyData(i)(4, 2)
        womanHos = womanHos + weeklyData(i)(4, 6)
    Next i

    ' Process dentist data
    For i = 0 To 6
        dentistPateints = dentistPateints + weeklyData(i)(5, 0)
        dentistTotal = dentistTotal + weeklyData(i)(5, 1)
        dentistHos = dentistHos + weeklyData(i)(5, 6)
    Next i

    ' Process ultrasound data
  
   For i = 0 To 6
        ultraPateints = ultraPateints + weeklyData(i)(6, 0)
        ultraTotal = ultraTotal + weeklyData(i)(6, 1)
        ultraYek = ultraYek + weeklyData(i)(6, 2)
        ultraRah = ultraRah + weeklyData(i)(6, 3)
        ultraShahim = ultraShahim + weeklyData(i)(6, 4)
        ultraWali = ultraWali + weeklyData(i)(6, 5)
        ultraHos = ultraHos + weeklyData(i)(6, 6)
    Next i

    ' Process lab data
    For i = 0 To 6
        labPateints = labPateints + weeklyData(i)(7, 0)
        labTotal = labTotal + weeklyData(i)(7, 1)
        labYek = labYek + weeklyData(i)(7, 2)
        labRah = labRah + weeklyData(i)(7, 3)
        labShahim = labShahim + weeklyData(i)(7, 4)
        labWali = labWali + weeklyData(i)(7, 5)
        labHos = labHos + weeklyData(i)(7, 6)
    Next i

    ' Process surgery data
   For i = 0 To 6
        surgPatients = surgPatients + weeklyData(i)(8, 0)
        surgTotal = surgTotal + weeklyData(i)(8, 1)
        surgHos = surgHos + weeklyData(i)(8, 6)
    Next i
    
    
    
   ' صدا زدن سب روتين براي ساخت جدول گزارش هفته در برگه گزارش هفته وار
    
    Call reportTable(empateint, emTotal, emHos, bedPatient, bedTotal, bedHos, inPatient, inTotal, inRah, _
    inShahim, inWali, inHos, childbirthPatients, childbirthTotal, childbirthHos, _
    childbirthHos, womanPatients, womanTotal, womanYek, womanHos, dentistPatients, _
    dentistTotal, dentistHos, ultraPateints, ultraTotal, ultraYek, ultraRah, ultraShahim, ultraWali, ultraHos, _
    labPateints, labTotal, labYek, labRah, labShahim, labWali, labHos, surgPatients, surgTotal, surgHos, exps, pharB, drugBuys, _
    startDate, endDate)
      

    
End Sub

