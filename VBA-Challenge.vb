{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub vba_challenge()\
\
'To loop across all worksheets\
For Each ws In Worksheets\
\
'Variables declaration:\
'For worksheet looping:\
Dim Worksheetname As String\
'For the first table:\
Dim ticker As String\
Dim lastrow, total_vol, openvalue, closevalue, yearchange, percentchng As Double\
'For the second table:\
Dim maxticker, minticker, maxvolticker As String\
Dim currentprcnt, maxprcnt, minprcnt, maxtotal As Double\
'For loops:\
Dim i, j, k As Integer\
\
Worksheetname = ws.Name\
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
maxprcnt = 0\
minprcnt = 0\
total_vol = 0\
j = 2\
\
'Creates columns:\
ws.Range("I1").Value = "Ticker"\
ws.Range("J1").Value = "Yearly Change"\
ws.Range("K1").Value = "Percent Change"\
ws.Range("L1").Value = "Total Stock Volume"\
ws.Range("O2").Value = "Greatest % Increase"\
ws.Range("O3").Value = "Greatest % Decrease"\
ws.Range("O4").Value = "Greatest Total Volume"\
ws.Range("P1").Value = "Ticker"\
ws.Range("Q1").Value = "Value"\
\
For i = 2 To lastrow\
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then\
        openvalue = ws.Cells(i, 3).Value\
    End If\
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then\
        'Just add to total volume\
        total_vol = total_vol + ws.Cells(i, 7).Value\
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
        'Calculate everything\
        total_vol = total_vol + ws.Cells(i, 7).Value\
        closevalue = ws.Cells(i, 6).Value\
        ticker = ws.Cells(i, 1).Value\
        yearchange = closevalue - openvalue\
        percentchng = yearchange / openvalue\
        'Fills the first table\
        ws.Cells(j, 9).Value = ticker\
        ws.Cells(j, 10).Value = yearchange\
        'Conditional formating for yearly change column:\
                If yearchange < 0 Then\
                    ws.Cells(j, 10).Interior.ColorIndex = 3\
                ElseIf yearchange > 0 Then\
                    ws.Cells(j, 10).Interior.ColorIndex = 4\
                End If\
        ws.Cells(j, 11).Value = FormatPercent(percentchng, 2)\
        'Calculates min and max percent change\
        If percentchng > maxprcnt Then\
            maxprcnt = percentchng\
            maxticker = ticker\
        ElseIf percentchng < minprcnt Then\
            minprcnt = percentchng\
            minticker = ticker\
        End If\
        ws.Cells(j, 12).Value = total_vol\
        'Calculates max volume\
        If total_vol > maxtotal Then\
            maxtotal = total_vol\
            maxvolticker = ticker\
        End If\
        total_vol = 0\
        j = j + 1\
    End If\
Next i\
\
'Fills the second table\
ws.Range("P2").Value = maxticker\
ws.Range("P3").Value = minticker\
ws.Range("Q2").Value = FormatPercent(maxprcnt, 2)\
ws.Range("Q3").Value = FormatPercent(minprcnt, 2)\
ws.Range("P4").Value = maxvolticker\
ws.Range("Q4").Value = maxtotal\
\
' Autofit to display data\
ws.Columns("A:Q").AutoFit\
\
Next ws\
End Sub\
}