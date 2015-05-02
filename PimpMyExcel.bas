Attribute VB_Name = "PimpMyExcel"
Sub PimpMyExcel()

    ' ScreenUpdating controls whether or not Excel refreshes the
    ' screen after performing certain operations. We don't want
    ' Excel to show any animations while running code. Setting this
    ' to false will reduce run time.
    Application.ScreenUpdating = False

    ' Sometimes Excel will throw an error relating to too many
    ' cell formats or fonts but we can just resume next to
    ' continue running the code.
    On Error Resume Next

    ' Make variables for the font color and the background color.
    Dim fontArrayGenerated() As Integer
    Dim backgroundArrayGenerated() As Integer

    ' Set font color of all cells to automatic. We do this because we
    ' can't go over the Excel maximum limit of 512 different fonts.
    Cells.Font.ColorIndex = xlAutomatic

    ' Clear the background color. We do this because in Excel 2000
    ' and later the maximum number of cell formats is approximately
    ' 4,000 different combinations and in Excel 2007 and higher the
    ' maximum number of cell formats is 64,000.
    Cells.Interior.Pattern = xlNone

    ' Loop through each cell in the worksheet by looping through each
    ' row and each column.
    For i = 1 To CalcLastRow()

        For j = 1 To CalcLastColumn()

            ' Generate a random background color from our GenerateBackgroundColor() function
            backgroundArrayGenerated = GenerateBackgroundColor()

            ' Set color of background of cell to this random color
            ActiveSheet.Cells(i, j).Interior.Color = RGB(backgroundArrayGenerated(0), backgroundArrayGenerated(1), backgroundArrayGenerated(2))

            ' Generate a random font color from our GenerateFontColor() function
            fontArrayGenerated = GenerateFontColor()

            ' Set color of font of cell to this random color
            ActiveSheet.Cells(i, j).Font.Color = RGB(fontArrayGenerated(0), fontArrayGenerated(1), fontArrayGenerated(2))

        Next j

    Next i

    ' We have to set ScreenUpdating to True before exiting or
    ' the user will not be able to make any changes to the
    ' worksheet.
    Application.ScreenUpdating = True

End Sub

Function GenerateFontColor() As Variant

    ' This function includes a prepopulated array of 100 different
    ' font colors. Excel has a limitation of 512 fonts that can be
    ' used at any given time. In order to prevent using more than the
    ' 512 fonts I generated these 100 random font colors. Running this
    ' program even repeatedly will therefore never uses more than the
    ' 100 font colors contained in this function. Randomly generating
    ' font colors for every cell also dramatically increases run time
    ' so to have hard coded font colors will decrease the runtime of
    ' the program.

    Dim fontColorIndex As Integer
    Dim fontArray(100, 2) As Integer
    Dim returnedFontArray(2) As Integer

    ' 100 randomly generated font colors
    fontArray(0, 0) = 82
    fontArray(0, 1) = 31
    fontArray(0, 2) = 238
    fontArray(1, 0) = 230
    fontArray(1, 1) = 84
    fontArray(1, 2) = 208
    fontArray(2, 0) = 67
    fontArray(2, 1) = 104
    fontArray(2, 2) = 247
    fontArray(3, 0) = 86
    fontArray(3, 1) = 65
    fontArray(3, 2) = 240
    fontArray(4, 0) = 41
    fontArray(4, 1) = 78
    fontArray(4, 2) = 28
    fontArray(5, 0) = 157
    fontArray(5, 1) = 196
    fontArray(5, 2) = 233
    fontArray(6, 0) = 32
    fontArray(6, 1) = 14
    fontArray(6, 2) = 160
    fontArray(7, 0) = 205
    fontArray(7, 1) = 216
    fontArray(7, 2) = 125
    fontArray(8, 0) = 86
    fontArray(8, 1) = 124
    fontArray(8, 2) = 98
    fontArray(9, 0) = 41
    fontArray(9, 1) = 119
    fontArray(9, 2) = 222
    fontArray(10, 0) = 239
    fontArray(10, 1) = 220
    fontArray(10, 2) = 20
    fontArray(11, 0) = 10
    fontArray(11, 1) = 177
    fontArray(11, 2) = 103
    fontArray(12, 0) = 198
    fontArray(12, 1) = 96
    fontArray(12, 2) = 67
    fontArray(13, 0) = 178
    fontArray(13, 1) = 167
    fontArray(13, 2) = 184
    fontArray(14, 0) = 4
    fontArray(14, 1) = 210
    fontArray(14, 2) = 162
    fontArray(15, 0) = 233
    fontArray(15, 1) = 139
    fontArray(15, 2) = 130
    fontArray(16, 0) = 38
    fontArray(16, 1) = 222
    fontArray(16, 2) = 174
    fontArray(17, 0) = 125
    fontArray(17, 1) = 11
    fontArray(17, 2) = 52
    fontArray(18, 0) = 90
    fontArray(18, 1) = 172
    fontArray(18, 2) = 14
    fontArray(19, 0) = 17
    fontArray(19, 1) = 98
    fontArray(19, 2) = 144
    fontArray(20, 0) = 90
    fontArray(20, 1) = 11
    fontArray(20, 2) = 20
    fontArray(21, 0) = 3
    fontArray(21, 1) = 36
    fontArray(21, 2) = 141
    fontArray(22, 0) = 65
    fontArray(22, 1) = 109
    fontArray(22, 2) = 157
    fontArray(23, 0) = 122
    fontArray(23, 1) = 252
    fontArray(23, 2) = 2
    fontArray(24, 0) = 42
    fontArray(24, 1) = 47
    fontArray(24, 2) = 118
    fontArray(25, 0) = 146
    fontArray(25, 1) = 96
    fontArray(25, 2) = 53
    fontArray(26, 0) = 86
    fontArray(26, 1) = 133
    fontArray(26, 2) = 204
    fontArray(27, 0) = 157
    fontArray(27, 1) = 93
    fontArray(27, 2) = 25
    fontArray(28, 0) = 124
    fontArray(28, 1) = 103
    fontArray(28, 2) = 241
    fontArray(29, 0) = 38
    fontArray(29, 1) = 250
    fontArray(29, 2) = 246
    fontArray(30, 0) = 135
    fontArray(30, 1) = 221
    fontArray(30, 2) = 178
    fontArray(31, 0) = 69
    fontArray(31, 1) = 65
    fontArray(31, 2) = 141
    fontArray(32, 0) = 17
    fontArray(32, 1) = 152
    fontArray(32, 2) = 153
    fontArray(33, 0) = 147
    fontArray(33, 1) = 204
    fontArray(33, 2) = 26
    fontArray(34, 0) = 182
    fontArray(34, 1) = 14
    fontArray(34, 2) = 18
    fontArray(35, 0) = 166
    fontArray(35, 1) = 35
    fontArray(35, 2) = 161
    fontArray(36, 0) = 233
    fontArray(36, 1) = 101
    fontArray(36, 2) = 218
    fontArray(37, 0) = 166
    fontArray(37, 1) = 43
    fontArray(37, 2) = 96
    fontArray(38, 0) = 137
    fontArray(38, 1) = 126
    fontArray(38, 2) = 199
    fontArray(39, 0) = 246
    fontArray(39, 1) = 59
    fontArray(39, 2) = 192
    fontArray(40, 0) = 205
    fontArray(40, 1) = 93
    fontArray(40, 2) = 230
    fontArray(41, 0) = 129
    fontArray(41, 1) = 65
    fontArray(41, 2) = 21
    fontArray(42, 0) = 235
    fontArray(42, 1) = 204
    fontArray(42, 2) = 100
    fontArray(43, 0) = 215
    fontArray(43, 1) = 217
    fontArray(43, 2) = 47
    fontArray(44, 0) = 165
    fontArray(44, 1) = 143
    fontArray(44, 2) = 253
    fontArray(45, 0) = 94
    fontArray(45, 1) = 241
    fontArray(45, 2) = 179
    fontArray(46, 0) = 176
    fontArray(46, 1) = 238
    fontArray(46, 2) = 85
    fontArray(47, 0) = 115
    fontArray(47, 1) = 11
    fontArray(47, 2) = 88
    fontArray(48, 0) = 139
    fontArray(48, 1) = 215
    fontArray(48, 2) = 251
    fontArray(49, 0) = 147
    fontArray(49, 1) = 41
    fontArray(49, 2) = 205
    fontArray(50, 0) = 178
    fontArray(50, 1) = 37
    fontArray(50, 2) = 36
    fontArray(51, 0) = 77
    fontArray(51, 1) = 92
    fontArray(51, 2) = 207
    fontArray(52, 0) = 196
    fontArray(52, 1) = 119
    fontArray(52, 2) = 252
    fontArray(53, 0) = 34
    fontArray(53, 1) = 174
    fontArray(53, 2) = 115
    fontArray(54, 0) = 20
    fontArray(54, 1) = 160
    fontArray(54, 2) = 201
    fontArray(55, 0) = 19
    fontArray(55, 1) = 242
    fontArray(55, 2) = 97
    fontArray(56, 0) = 131
    fontArray(56, 1) = 27
    fontArray(56, 2) = 182
    fontArray(57, 0) = 128
    fontArray(57, 1) = 112
    fontArray(57, 2) = 225
    fontArray(58, 0) = 21
    fontArray(58, 1) = 32
    fontArray(58, 2) = 83
    fontArray(59, 0) = 71
    fontArray(59, 1) = 191
    fontArray(59, 2) = 99
    fontArray(60, 0) = 167
    fontArray(60, 1) = 223
    fontArray(60, 2) = 45
    fontArray(61, 0) = 20
    fontArray(61, 1) = 85
    fontArray(61, 2) = 129
    fontArray(62, 0) = 96
    fontArray(62, 1) = 159
    fontArray(62, 2) = 36
    fontArray(63, 0) = 81
    fontArray(63, 1) = 92
    fontArray(63, 2) = 3
    fontArray(64, 0) = 232
    fontArray(64, 1) = 136
    fontArray(64, 2) = 76
    fontArray(65, 0) = 10
    fontArray(65, 1) = 215
    fontArray(65, 2) = 132
    fontArray(66, 0) = 38
    fontArray(66, 1) = 17
    fontArray(66, 2) = 172
    fontArray(67, 0) = 122
    fontArray(67, 1) = 101
    fontArray(67, 2) = 140
    fontArray(68, 0) = 194
    fontArray(68, 1) = 34
    fontArray(68, 2) = 87
    fontArray(69, 0) = 195
    fontArray(69, 1) = 170
    fontArray(69, 2) = 102
    fontArray(70, 0) = 160
    fontArray(70, 1) = 170
    fontArray(70, 2) = 161
    fontArray(71, 0) = 174
    fontArray(71, 1) = 221
    fontArray(71, 2) = 126
    fontArray(72, 0) = 20
    fontArray(72, 1) = 181
    fontArray(72, 2) = 31
    fontArray(73, 0) = 246
    fontArray(73, 1) = 34
    fontArray(73, 2) = 84
    fontArray(74, 0) = 119
    fontArray(74, 1) = 245
    fontArray(74, 2) = 236
    fontArray(75, 0) = 83
    fontArray(75, 1) = 58
    fontArray(75, 2) = 91
    fontArray(76, 0) = 148
    fontArray(76, 1) = 112
    fontArray(76, 2) = 134
    fontArray(77, 0) = 201
    fontArray(77, 1) = 243
    fontArray(77, 2) = 174
    fontArray(78, 0) = 62
    fontArray(78, 1) = 138
    fontArray(78, 2) = 129
    fontArray(79, 0) = 43
    fontArray(79, 1) = 148
    fontArray(79, 2) = 80
    fontArray(80, 0) = 16
    fontArray(80, 1) = 206
    fontArray(80, 2) = 62
    fontArray(81, 0) = 106
    fontArray(81, 1) = 169
    fontArray(81, 2) = 104
    fontArray(82, 0) = 175
    fontArray(82, 1) = 142
    fontArray(82, 2) = 197
    fontArray(83, 0) = 115
    fontArray(83, 1) = 190
    fontArray(83, 2) = 106
    fontArray(84, 0) = 43
    fontArray(84, 1) = 130
    fontArray(84, 2) = 36
    fontArray(85, 0) = 6
    fontArray(85, 1) = 255
    fontArray(85, 2) = 6
    fontArray(86, 0) = 226
    fontArray(86, 1) = 97
    fontArray(86, 2) = 40
    fontArray(87, 0) = 179
    fontArray(87, 1) = 149
    fontArray(87, 2) = 47
    fontArray(88, 0) = 74
    fontArray(88, 1) = 102
    fontArray(88, 2) = 116
    fontArray(89, 0) = 105
    fontArray(89, 1) = 104
    fontArray(89, 2) = 59
    fontArray(90, 0) = 199
    fontArray(90, 1) = 106
    fontArray(90, 2) = 103
    fontArray(91, 0) = 132
    fontArray(91, 1) = 92
    fontArray(91, 2) = 109
    fontArray(92, 0) = 122
    fontArray(92, 1) = 168
    fontArray(92, 2) = 31
    fontArray(93, 0) = 166
    fontArray(93, 1) = 225
    fontArray(93, 2) = 107
    fontArray(94, 0) = 123
    fontArray(94, 1) = 146
    fontArray(94, 2) = 11
    fontArray(95, 0) = 45
    fontArray(95, 1) = 58
    fontArray(95, 2) = 197
    fontArray(96, 0) = 33
    fontArray(96, 1) = 128
    fontArray(96, 2) = 14
    fontArray(97, 0) = 32
    fontArray(97, 1) = 213
    fontArray(97, 2) = 8
    fontArray(98, 0) = 187
    fontArray(98, 1) = 236
    fontArray(98, 2) = 212
    fontArray(99, 0) = 37
    fontArray(99, 1) = 202
    fontArray(99, 2) = 254

    ' Randomly pick one of the 100 font colors to return
    fontColorIndex = Int((100) * Rnd)

    ' Assign the RGB values of the font color to the array
    returnedFontArray(0) = fontArray(fontColorIndex, 0)
    returnedFontArray(1) = fontArray(fontColorIndex, 1)
    returnedFontArray(2) = fontArray(fontColorIndex, 2)

    ' Return the font color array to the main program
    GenerateFontColor = returnedFontArray

End Function

Function GenerateBackgroundColor() As Variant

    ' This function includes a prepopulated array of 400 different
    ' background colors. Excel has a limitation of 4,000 cell formats
    ' in Excel 2000 and up and a limit of 64,000 cell formats in Excel
    ' 2007 and higher. In order to prevent using more than the 4,000 or
    ' 64,000 cell formats I generated these 400 random background colors.
    ' Running this program even repeatedly will therefore never uses more
    ' than the 400 background colors contained in this function. Randomly
    ' generating background colors for every cell also dramatically increases
    ' run time so to have hard coded background colors will decrease the
    ' runtime of the program.

    Dim backgroundColorIndex As Integer
    Dim backgroundArray(400, 2) As Integer
    Dim returnedBackgroundArray(2) As Integer

    ' 400 randomly generated background colors

    backgroundArray(0, 0) = 160
    backgroundArray(0, 1) = 14
    backgroundArray(0, 2) = 212
    backgroundArray(1, 0) = 105
    backgroundArray(1, 1) = 21
    backgroundArray(1, 2) = 171
    backgroundArray(2, 0) = 201
    backgroundArray(2, 1) = 217
    backgroundArray(2, 2) = 197
    backgroundArray(3, 0) = 89
    backgroundArray(3, 1) = 131
    backgroundArray(3, 2) = 42
    backgroundArray(4, 0) = 2
    backgroundArray(4, 1) = 216
    backgroundArray(4, 2) = 110
    backgroundArray(5, 0) = 13
    backgroundArray(5, 1) = 112
    backgroundArray(5, 2) = 119
    backgroundArray(6, 0) = 248
    backgroundArray(6, 1) = 165
    backgroundArray(6, 2) = 163
    backgroundArray(7, 0) = 226
    backgroundArray(7, 1) = 112
    backgroundArray(7, 2) = 166
    backgroundArray(8, 0) = 79
    backgroundArray(8, 1) = 109
    backgroundArray(8, 2) = 46
    backgroundArray(9, 0) = 239
    backgroundArray(9, 1) = 93
    backgroundArray(9, 2) = 122
    backgroundArray(10, 0) = 141
    backgroundArray(10, 1) = 104
    backgroundArray(10, 2) = 46
    backgroundArray(11, 0) = 79
    backgroundArray(11, 1) = 75
    backgroundArray(11, 2) = 236
    backgroundArray(12, 0) = 208
    backgroundArray(12, 1) = 114
    backgroundArray(12, 2) = 221
    backgroundArray(13, 0) = 239
    backgroundArray(13, 1) = 144
    backgroundArray(13, 2) = 204
    backgroundArray(14, 0) = 19
    backgroundArray(14, 1) = 35
    backgroundArray(14, 2) = 236
    backgroundArray(15, 0) = 134
    backgroundArray(15, 1) = 1
    backgroundArray(15, 2) = 30
    backgroundArray(16, 0) = 190
    backgroundArray(16, 1) = 166
    backgroundArray(16, 2) = 187
    backgroundArray(17, 0) = 247
    backgroundArray(17, 1) = 241
    backgroundArray(17, 2) = 153
    backgroundArray(18, 0) = 156
    backgroundArray(18, 1) = 79
    backgroundArray(18, 2) = 189
    backgroundArray(19, 0) = 205
    backgroundArray(19, 1) = 13
    backgroundArray(19, 2) = 132
    backgroundArray(20, 0) = 3
    backgroundArray(20, 1) = 217
    backgroundArray(20, 2) = 214
    backgroundArray(21, 0) = 65
    backgroundArray(21, 1) = 44
    backgroundArray(21, 2) = 210
    backgroundArray(22, 0) = 115
    backgroundArray(22, 1) = 228
    backgroundArray(22, 2) = 61
    backgroundArray(23, 0) = 91
    backgroundArray(23, 1) = 163
    backgroundArray(23, 2) = 97
    backgroundArray(24, 0) = 212
    backgroundArray(24, 1) = 88
    backgroundArray(24, 2) = 37
    backgroundArray(25, 0) = 119
    backgroundArray(25, 1) = 110
    backgroundArray(25, 2) = 250
    backgroundArray(26, 0) = 108
    backgroundArray(26, 1) = 93
    backgroundArray(26, 2) = 78
    backgroundArray(27, 0) = 78
    backgroundArray(27, 1) = 107
    backgroundArray(27, 2) = 148
    backgroundArray(28, 0) = 127
    backgroundArray(28, 1) = 165
    backgroundArray(28, 2) = 93
    backgroundArray(29, 0) = 127
    backgroundArray(29, 1) = 134
    backgroundArray(29, 2) = 20
    backgroundArray(30, 0) = 164
    backgroundArray(30, 1) = 63
    backgroundArray(30, 2) = 230
    backgroundArray(31, 0) = 40
    backgroundArray(31, 1) = 184
    backgroundArray(31, 2) = 246
    backgroundArray(32, 0) = 109
    backgroundArray(32, 1) = 214
    backgroundArray(32, 2) = 102
    backgroundArray(33, 0) = 34
    backgroundArray(33, 1) = 164
    backgroundArray(33, 2) = 167
    backgroundArray(34, 0) = 81
    backgroundArray(34, 1) = 203
    backgroundArray(34, 2) = 51
    backgroundArray(35, 0) = 6
    backgroundArray(35, 1) = 126
    backgroundArray(35, 2) = 145
    backgroundArray(36, 0) = 179
    backgroundArray(36, 1) = 211
    backgroundArray(36, 2) = 205
    backgroundArray(37, 0) = 97
    backgroundArray(37, 1) = 164
    backgroundArray(37, 2) = 4
    backgroundArray(38, 0) = 235
    backgroundArray(38, 1) = 168
    backgroundArray(38, 2) = 96
    backgroundArray(39, 0) = 28
    backgroundArray(39, 1) = 22
    backgroundArray(39, 2) = 183
    backgroundArray(40, 0) = 118
    backgroundArray(40, 1) = 21
    backgroundArray(40, 2) = 41
    backgroundArray(41, 0) = 52
    backgroundArray(41, 1) = 157
    backgroundArray(41, 2) = 62
    backgroundArray(42, 0) = 173
    backgroundArray(42, 1) = 124
    backgroundArray(42, 2) = 51
    backgroundArray(43, 0) = 46
    backgroundArray(43, 1) = 206
    backgroundArray(43, 2) = 30
    backgroundArray(44, 0) = 245
    backgroundArray(44, 1) = 218
    backgroundArray(44, 2) = 182
    backgroundArray(45, 0) = 42
    backgroundArray(45, 1) = 90
    backgroundArray(45, 2) = 123
    backgroundArray(46, 0) = 21
    backgroundArray(46, 1) = 192
    backgroundArray(46, 2) = 197
    backgroundArray(47, 0) = 251
    backgroundArray(47, 1) = 231
    backgroundArray(47, 2) = 160
    backgroundArray(48, 0) = 121
    backgroundArray(48, 1) = 135
    backgroundArray(48, 2) = 255
    backgroundArray(49, 0) = 202
    backgroundArray(49, 1) = 120
    backgroundArray(49, 2) = 19
    backgroundArray(50, 0) = 170
    backgroundArray(50, 1) = 51
    backgroundArray(50, 2) = 196
    backgroundArray(51, 0) = 192
    backgroundArray(51, 1) = 185
    backgroundArray(51, 2) = 101
    backgroundArray(52, 0) = 11
    backgroundArray(52, 1) = 2
    backgroundArray(52, 2) = 173
    backgroundArray(53, 0) = 208
    backgroundArray(53, 1) = 55
    backgroundArray(53, 2) = 35
    backgroundArray(54, 0) = 154
    backgroundArray(54, 1) = 23
    backgroundArray(54, 2) = 140
    backgroundArray(55, 0) = 72
    backgroundArray(55, 1) = 112
    backgroundArray(55, 2) = 252
    backgroundArray(56, 0) = 139
    backgroundArray(56, 1) = 42
    backgroundArray(56, 2) = 102
    backgroundArray(57, 0) = 99
    backgroundArray(57, 1) = 53
    backgroundArray(57, 2) = 141
    backgroundArray(58, 0) = 147
    backgroundArray(58, 1) = 111
    backgroundArray(58, 2) = 48
    backgroundArray(59, 0) = 238
    backgroundArray(59, 1) = 120
    backgroundArray(59, 2) = 101
    backgroundArray(60, 0) = 123
    backgroundArray(60, 1) = 131
    backgroundArray(60, 2) = 191
    backgroundArray(61, 0) = 98
    backgroundArray(61, 1) = 158
    backgroundArray(61, 2) = 87
    backgroundArray(62, 0) = 227
    backgroundArray(62, 1) = 89
    backgroundArray(62, 2) = 237
    backgroundArray(63, 0) = 251
    backgroundArray(63, 1) = 42
    backgroundArray(63, 2) = 125
    backgroundArray(64, 0) = 69
    backgroundArray(64, 1) = 80
    backgroundArray(64, 2) = 29
    backgroundArray(65, 0) = 151
    backgroundArray(65, 1) = 53
    backgroundArray(65, 2) = 204
    backgroundArray(66, 0) = 254
    backgroundArray(66, 1) = 74
    backgroundArray(66, 2) = 215
    backgroundArray(67, 0) = 113
    backgroundArray(67, 1) = 228
    backgroundArray(67, 2) = 58
    backgroundArray(68, 0) = 166
    backgroundArray(68, 1) = 38
    backgroundArray(68, 2) = 127
    backgroundArray(69, 0) = 244
    backgroundArray(69, 1) = 0
    backgroundArray(69, 2) = 20
    backgroundArray(70, 0) = 229
    backgroundArray(70, 1) = 203
    backgroundArray(70, 2) = 198
    backgroundArray(71, 0) = 9
    backgroundArray(71, 1) = 203
    backgroundArray(71, 2) = 160
    backgroundArray(72, 0) = 118
    backgroundArray(72, 1) = 136
    backgroundArray(72, 2) = 56
    backgroundArray(73, 0) = 141
    backgroundArray(73, 1) = 107
    backgroundArray(73, 2) = 112
    backgroundArray(74, 0) = 28
    backgroundArray(74, 1) = 36
    backgroundArray(74, 2) = 213
    backgroundArray(75, 0) = 135
    backgroundArray(75, 1) = 222
    backgroundArray(75, 2) = 105
    backgroundArray(76, 0) = 255
    backgroundArray(76, 1) = 74
    backgroundArray(76, 2) = 179
    backgroundArray(77, 0) = 99
    backgroundArray(77, 1) = 252
    backgroundArray(77, 2) = 37
    backgroundArray(78, 0) = 123
    backgroundArray(78, 1) = 142
    backgroundArray(78, 2) = 12
    backgroundArray(79, 0) = 143
    backgroundArray(79, 1) = 181
    backgroundArray(79, 2) = 96
    backgroundArray(80, 0) = 0
    backgroundArray(80, 1) = 73
    backgroundArray(80, 2) = 197
    backgroundArray(81, 0) = 191
    backgroundArray(81, 1) = 240
    backgroundArray(81, 2) = 246
    backgroundArray(82, 0) = 219
    backgroundArray(82, 1) = 228
    backgroundArray(82, 2) = 32
    backgroundArray(83, 0) = 225
    backgroundArray(83, 1) = 38
    backgroundArray(83, 2) = 219
    backgroundArray(84, 0) = 183
    backgroundArray(84, 1) = 87
    backgroundArray(84, 2) = 96
    backgroundArray(85, 0) = 95
    backgroundArray(85, 1) = 238
    backgroundArray(85, 2) = 210
    backgroundArray(86, 0) = 170
    backgroundArray(86, 1) = 50
    backgroundArray(86, 2) = 158
    backgroundArray(87, 0) = 11
    backgroundArray(87, 1) = 23
    backgroundArray(87, 2) = 19
    backgroundArray(88, 0) = 98
    backgroundArray(88, 1) = 14
    backgroundArray(88, 2) = 189
    backgroundArray(89, 0) = 116
    backgroundArray(89, 1) = 152
    backgroundArray(89, 2) = 3
    backgroundArray(90, 0) = 15
    backgroundArray(90, 1) = 132
    backgroundArray(90, 2) = 153
    backgroundArray(91, 0) = 168
    backgroundArray(91, 1) = 72
    backgroundArray(91, 2) = 179
    backgroundArray(92, 0) = 178
    backgroundArray(92, 1) = 253
    backgroundArray(92, 2) = 247
    backgroundArray(93, 0) = 39
    backgroundArray(93, 1) = 234
    backgroundArray(93, 2) = 5
    backgroundArray(94, 0) = 46
    backgroundArray(94, 1) = 232
    backgroundArray(94, 2) = 147
    backgroundArray(95, 0) = 195
    backgroundArray(95, 1) = 188
    backgroundArray(95, 2) = 21
    backgroundArray(96, 0) = 156
    backgroundArray(96, 1) = 81
    backgroundArray(96, 2) = 74
    backgroundArray(97, 0) = 87
    backgroundArray(97, 1) = 21
    backgroundArray(97, 2) = 97
    backgroundArray(98, 0) = 116
    backgroundArray(98, 1) = 204
    backgroundArray(98, 2) = 229
    backgroundArray(99, 0) = 133
    backgroundArray(99, 1) = 83
    backgroundArray(99, 2) = 209
    backgroundArray(100, 0) = 98
    backgroundArray(100, 1) = 56
    backgroundArray(100, 2) = 1
    backgroundArray(101, 0) = 45
    backgroundArray(101, 1) = 134
    backgroundArray(101, 2) = 116
    backgroundArray(102, 0) = 208
    backgroundArray(102, 1) = 192
    backgroundArray(102, 2) = 123
    backgroundArray(103, 0) = 5
    backgroundArray(103, 1) = 136
    backgroundArray(103, 2) = 176
    backgroundArray(104, 0) = 127
    backgroundArray(104, 1) = 195
    backgroundArray(104, 2) = 129
    backgroundArray(105, 0) = 250
    backgroundArray(105, 1) = 122
    backgroundArray(105, 2) = 34
    backgroundArray(106, 0) = 104
    backgroundArray(106, 1) = 117
    backgroundArray(106, 2) = 220
    backgroundArray(107, 0) = 80
    backgroundArray(107, 1) = 82
    backgroundArray(107, 2) = 250
    backgroundArray(108, 0) = 207
    backgroundArray(108, 1) = 193
    backgroundArray(108, 2) = 113
    backgroundArray(109, 0) = 87
    backgroundArray(109, 1) = 178
    backgroundArray(109, 2) = 154
    backgroundArray(110, 0) = 101
    backgroundArray(110, 1) = 79
    backgroundArray(110, 2) = 229
    backgroundArray(111, 0) = 191
    backgroundArray(111, 1) = 16
    backgroundArray(111, 2) = 87
    backgroundArray(112, 0) = 117
    backgroundArray(112, 1) = 90
    backgroundArray(112, 2) = 241
    backgroundArray(113, 0) = 94
    backgroundArray(113, 1) = 64
    backgroundArray(113, 2) = 213
    backgroundArray(114, 0) = 138
    backgroundArray(114, 1) = 128
    backgroundArray(114, 2) = 100
    backgroundArray(115, 0) = 216
    backgroundArray(115, 1) = 88
    backgroundArray(115, 2) = 225
    backgroundArray(116, 0) = 212
    backgroundArray(116, 1) = 207
    backgroundArray(116, 2) = 203
    backgroundArray(117, 0) = 25
    backgroundArray(117, 1) = 177
    backgroundArray(117, 2) = 221
    backgroundArray(118, 0) = 77
    backgroundArray(118, 1) = 235
    backgroundArray(118, 2) = 164
    backgroundArray(119, 0) = 222
    backgroundArray(119, 1) = 7
    backgroundArray(119, 2) = 130
    backgroundArray(120, 0) = 4
    backgroundArray(120, 1) = 190
    backgroundArray(120, 2) = 122
    backgroundArray(121, 0) = 99
    backgroundArray(121, 1) = 149
    backgroundArray(121, 2) = 128
    backgroundArray(122, 0) = 119
    backgroundArray(122, 1) = 221
    backgroundArray(122, 2) = 88
    backgroundArray(123, 0) = 186
    backgroundArray(123, 1) = 170
    backgroundArray(123, 2) = 13
    backgroundArray(124, 0) = 94
    backgroundArray(124, 1) = 12
    backgroundArray(124, 2) = 139
    backgroundArray(125, 0) = 132
    backgroundArray(125, 1) = 17
    backgroundArray(125, 2) = 192
    backgroundArray(126, 0) = 205
    backgroundArray(126, 1) = 32
    backgroundArray(126, 2) = 17
    backgroundArray(127, 0) = 150
    backgroundArray(127, 1) = 72
    backgroundArray(127, 2) = 151
    backgroundArray(128, 0) = 37
    backgroundArray(128, 1) = 172
    backgroundArray(128, 2) = 232
    backgroundArray(129, 0) = 168
    backgroundArray(129, 1) = 156
    backgroundArray(129, 2) = 61
    backgroundArray(130, 0) = 37
    backgroundArray(130, 1) = 93
    backgroundArray(130, 2) = 57
    backgroundArray(131, 0) = 199
    backgroundArray(131, 1) = 101
    backgroundArray(131, 2) = 161
    backgroundArray(132, 0) = 144
    backgroundArray(132, 1) = 231
    backgroundArray(132, 2) = 123
    backgroundArray(133, 0) = 51
    backgroundArray(133, 1) = 187
    backgroundArray(133, 2) = 212
    backgroundArray(134, 0) = 168
    backgroundArray(134, 1) = 145
    backgroundArray(134, 2) = 180
    backgroundArray(135, 0) = 134
    backgroundArray(135, 1) = 192
    backgroundArray(135, 2) = 235
    backgroundArray(136, 0) = 17
    backgroundArray(136, 1) = 167
    backgroundArray(136, 2) = 71
    backgroundArray(137, 0) = 16
    backgroundArray(137, 1) = 165
    backgroundArray(137, 2) = 11
    backgroundArray(138, 0) = 205
    backgroundArray(138, 1) = 118
    backgroundArray(138, 2) = 240
    backgroundArray(139, 0) = 3
    backgroundArray(139, 1) = 175
    backgroundArray(139, 2) = 23
    backgroundArray(140, 0) = 25
    backgroundArray(140, 1) = 80
    backgroundArray(140, 2) = 7
    backgroundArray(141, 0) = 106
    backgroundArray(141, 1) = 85
    backgroundArray(141, 2) = 223
    backgroundArray(142, 0) = 51
    backgroundArray(142, 1) = 74
    backgroundArray(142, 2) = 39
    backgroundArray(143, 0) = 248
    backgroundArray(143, 1) = 248
    backgroundArray(143, 2) = 252
    backgroundArray(144, 0) = 194
    backgroundArray(144, 1) = 5
    backgroundArray(144, 2) = 167
    backgroundArray(145, 0) = 252
    backgroundArray(145, 1) = 131
    backgroundArray(145, 2) = 254
    backgroundArray(146, 0) = 12
    backgroundArray(146, 1) = 93
    backgroundArray(146, 2) = 177
    backgroundArray(147, 0) = 203
    backgroundArray(147, 1) = 120
    backgroundArray(147, 2) = 216
    backgroundArray(148, 0) = 73
    backgroundArray(148, 1) = 29
    backgroundArray(148, 2) = 38
    backgroundArray(149, 0) = 139
    backgroundArray(149, 1) = 204
    backgroundArray(149, 2) = 207
    backgroundArray(150, 0) = 93
    backgroundArray(150, 1) = 193
    backgroundArray(150, 2) = 47
    backgroundArray(151, 0) = 76
    backgroundArray(151, 1) = 120
    backgroundArray(151, 2) = 22
    backgroundArray(152, 0) = 103
    backgroundArray(152, 1) = 177
    backgroundArray(152, 2) = 1
    backgroundArray(153, 0) = 179
    backgroundArray(153, 1) = 254
    backgroundArray(153, 2) = 231
    backgroundArray(154, 0) = 149
    backgroundArray(154, 1) = 127
    backgroundArray(154, 2) = 209
    backgroundArray(155, 0) = 33
    backgroundArray(155, 1) = 123
    backgroundArray(155, 2) = 92
    backgroundArray(156, 0) = 11
    backgroundArray(156, 1) = 159
    backgroundArray(156, 2) = 94
    backgroundArray(157, 0) = 136
    backgroundArray(157, 1) = 12
    backgroundArray(157, 2) = 134
    backgroundArray(158, 0) = 107
    backgroundArray(158, 1) = 229
    backgroundArray(158, 2) = 91
    backgroundArray(159, 0) = 12
    backgroundArray(159, 1) = 160
    backgroundArray(159, 2) = 59
    backgroundArray(160, 0) = 241
    backgroundArray(160, 1) = 227
    backgroundArray(160, 2) = 59
    backgroundArray(161, 0) = 12
    backgroundArray(161, 1) = 202
    backgroundArray(161, 2) = 46
    backgroundArray(162, 0) = 44
    backgroundArray(162, 1) = 92
    backgroundArray(162, 2) = 245
    backgroundArray(163, 0) = 30
    backgroundArray(163, 1) = 240
    backgroundArray(163, 2) = 100
    backgroundArray(164, 0) = 171
    backgroundArray(164, 1) = 45
    backgroundArray(164, 2) = 25
    backgroundArray(165, 0) = 59
    backgroundArray(165, 1) = 212
    backgroundArray(165, 2) = 39
    backgroundArray(166, 0) = 132
    backgroundArray(166, 1) = 144
    backgroundArray(166, 2) = 200
    backgroundArray(167, 0) = 223
    backgroundArray(167, 1) = 148
    backgroundArray(167, 2) = 19
    backgroundArray(168, 0) = 221
    backgroundArray(168, 1) = 29
    backgroundArray(168, 2) = 249
    backgroundArray(169, 0) = 230
    backgroundArray(169, 1) = 111
    backgroundArray(169, 2) = 209
    backgroundArray(170, 0) = 30
    backgroundArray(170, 1) = 2
    backgroundArray(170, 2) = 253
    backgroundArray(171, 0) = 137
    backgroundArray(171, 1) = 150
    backgroundArray(171, 2) = 213
    backgroundArray(172, 0) = 163
    backgroundArray(172, 1) = 83
    backgroundArray(172, 2) = 101
    backgroundArray(173, 0) = 199
    backgroundArray(173, 1) = 64
    backgroundArray(173, 2) = 245
    backgroundArray(174, 0) = 95
    backgroundArray(174, 1) = 76
    backgroundArray(174, 2) = 112
    backgroundArray(175, 0) = 11
    backgroundArray(175, 1) = 102
    backgroundArray(175, 2) = 236
    backgroundArray(176, 0) = 69
    backgroundArray(176, 1) = 58
    backgroundArray(176, 2) = 154
    backgroundArray(177, 0) = 103
    backgroundArray(177, 1) = 114
    backgroundArray(177, 2) = 138
    backgroundArray(178, 0) = 233
    backgroundArray(178, 1) = 137
    backgroundArray(178, 2) = 154
    backgroundArray(179, 0) = 23
    backgroundArray(179, 1) = 115
    backgroundArray(179, 2) = 171
    backgroundArray(180, 0) = 107
    backgroundArray(180, 1) = 35
    backgroundArray(180, 2) = 3
    backgroundArray(181, 0) = 173
    backgroundArray(181, 1) = 92
    backgroundArray(181, 2) = 165
    backgroundArray(182, 0) = 93
    backgroundArray(182, 1) = 200
    backgroundArray(182, 2) = 115
    backgroundArray(183, 0) = 161
    backgroundArray(183, 1) = 245
    backgroundArray(183, 2) = 71
    backgroundArray(184, 0) = 166
    backgroundArray(184, 1) = 162
    backgroundArray(184, 2) = 124
    backgroundArray(185, 0) = 99
    backgroundArray(185, 1) = 143
    backgroundArray(185, 2) = 92
    backgroundArray(186, 0) = 141
    backgroundArray(186, 1) = 118
    backgroundArray(186, 2) = 75
    backgroundArray(187, 0) = 43
    backgroundArray(187, 1) = 40
    backgroundArray(187, 2) = 24
    backgroundArray(188, 0) = 219
    backgroundArray(188, 1) = 99
    backgroundArray(188, 2) = 99
    backgroundArray(189, 0) = 185
    backgroundArray(189, 1) = 46
    backgroundArray(189, 2) = 106
    backgroundArray(190, 0) = 13
    backgroundArray(190, 1) = 249
    backgroundArray(190, 2) = 193
    backgroundArray(191, 0) = 169
    backgroundArray(191, 1) = 170
    backgroundArray(191, 2) = 225
    backgroundArray(192, 0) = 106
    backgroundArray(192, 1) = 106
    backgroundArray(192, 2) = 193
    backgroundArray(193, 0) = 188
    backgroundArray(193, 1) = 43
    backgroundArray(193, 2) = 179
    backgroundArray(194, 0) = 37
    backgroundArray(194, 1) = 185
    backgroundArray(194, 2) = 102
    backgroundArray(195, 0) = 119
    backgroundArray(195, 1) = 106
    backgroundArray(195, 2) = 155
    backgroundArray(196, 0) = 212
    backgroundArray(196, 1) = 27
    backgroundArray(196, 2) = 15
    backgroundArray(197, 0) = 10
    backgroundArray(197, 1) = 10
    backgroundArray(197, 2) = 242
    backgroundArray(198, 0) = 80
    backgroundArray(198, 1) = 167
    backgroundArray(198, 2) = 105
    backgroundArray(199, 0) = 199
    backgroundArray(199, 1) = 225
    backgroundArray(199, 2) = 77
    backgroundArray(200, 0) = 13
    backgroundArray(200, 1) = 212
    backgroundArray(200, 2) = 223
    backgroundArray(201, 0) = 187
    backgroundArray(201, 1) = 9
    backgroundArray(201, 2) = 230
    backgroundArray(202, 0) = 114
    backgroundArray(202, 1) = 67
    backgroundArray(202, 2) = 182
    backgroundArray(203, 0) = 7
    backgroundArray(203, 1) = 81
    backgroundArray(203, 2) = 193
    backgroundArray(204, 0) = 38
    backgroundArray(204, 1) = 103
    backgroundArray(204, 2) = 252
    backgroundArray(205, 0) = 251
    backgroundArray(205, 1) = 40
    backgroundArray(205, 2) = 142
    backgroundArray(206, 0) = 172
    backgroundArray(206, 1) = 187
    backgroundArray(206, 2) = 85
    backgroundArray(207, 0) = 216
    backgroundArray(207, 1) = 84
    backgroundArray(207, 2) = 63
    backgroundArray(208, 0) = 62
    backgroundArray(208, 1) = 127
    backgroundArray(208, 2) = 247
    backgroundArray(209, 0) = 59
    backgroundArray(209, 1) = 43
    backgroundArray(209, 2) = 128
    backgroundArray(210, 0) = 91
    backgroundArray(210, 1) = 130
    backgroundArray(210, 2) = 60
    backgroundArray(211, 0) = 245
    backgroundArray(211, 1) = 174
    backgroundArray(211, 2) = 248
    backgroundArray(212, 0) = 224
    backgroundArray(212, 1) = 137
    backgroundArray(212, 2) = 173
    backgroundArray(213, 0) = 196
    backgroundArray(213, 1) = 169
    backgroundArray(213, 2) = 148
    backgroundArray(214, 0) = 110
    backgroundArray(214, 1) = 92
    backgroundArray(214, 2) = 130
    backgroundArray(215, 0) = 158
    backgroundArray(215, 1) = 204
    backgroundArray(215, 2) = 4
    backgroundArray(216, 0) = 39
    backgroundArray(216, 1) = 19
    backgroundArray(216, 2) = 198
    backgroundArray(217, 0) = 34
    backgroundArray(217, 1) = 145
    backgroundArray(217, 2) = 129
    backgroundArray(218, 0) = 190
    backgroundArray(218, 1) = 178
    backgroundArray(218, 2) = 70
    backgroundArray(219, 0) = 73
    backgroundArray(219, 1) = 240
    backgroundArray(219, 2) = 207
    backgroundArray(220, 0) = 11
    backgroundArray(220, 1) = 67
    backgroundArray(220, 2) = 95
    backgroundArray(221, 0) = 215
    backgroundArray(221, 1) = 58
    backgroundArray(221, 2) = 179
    backgroundArray(222, 0) = 47
    backgroundArray(222, 1) = 16
    backgroundArray(222, 2) = 151
    backgroundArray(223, 0) = 227
    backgroundArray(223, 1) = 141
    backgroundArray(223, 2) = 248
    backgroundArray(224, 0) = 154
    backgroundArray(224, 1) = 237
    backgroundArray(224, 2) = 110
    backgroundArray(225, 0) = 219
    backgroundArray(225, 1) = 34
    backgroundArray(225, 2) = 138
    backgroundArray(226, 0) = 179
    backgroundArray(226, 1) = 117
    backgroundArray(226, 2) = 208
    backgroundArray(227, 0) = 143
    backgroundArray(227, 1) = 154
    backgroundArray(227, 2) = 218
    backgroundArray(228, 0) = 90
    backgroundArray(228, 1) = 88
    backgroundArray(228, 2) = 234
    backgroundArray(229, 0) = 181
    backgroundArray(229, 1) = 220
    backgroundArray(229, 2) = 148
    backgroundArray(230, 0) = 61
    backgroundArray(230, 1) = 41
    backgroundArray(230, 2) = 124
    backgroundArray(231, 0) = 237
    backgroundArray(231, 1) = 44
    backgroundArray(231, 2) = 48
    backgroundArray(232, 0) = 242
    backgroundArray(232, 1) = 153
    backgroundArray(232, 2) = 198
    backgroundArray(233, 0) = 241
    backgroundArray(233, 1) = 243
    backgroundArray(233, 2) = 59
    backgroundArray(234, 0) = 151
    backgroundArray(234, 1) = 93
    backgroundArray(234, 2) = 131
    backgroundArray(235, 0) = 90
    backgroundArray(235, 1) = 30
    backgroundArray(235, 2) = 33
    backgroundArray(236, 0) = 203
    backgroundArray(236, 1) = 112
    backgroundArray(236, 2) = 213
    backgroundArray(237, 0) = 227
    backgroundArray(237, 1) = 125
    backgroundArray(237, 2) = 145
    backgroundArray(238, 0) = 15
    backgroundArray(238, 1) = 154
    backgroundArray(238, 2) = 87
    backgroundArray(239, 0) = 46
    backgroundArray(239, 1) = 126
    backgroundArray(239, 2) = 31
    backgroundArray(240, 0) = 103
    backgroundArray(240, 1) = 148
    backgroundArray(240, 2) = 245
    backgroundArray(241, 0) = 12
    backgroundArray(241, 1) = 249
    backgroundArray(241, 2) = 180
    backgroundArray(242, 0) = 53
    backgroundArray(242, 1) = 206
    backgroundArray(242, 2) = 123
    backgroundArray(243, 0) = 189
    backgroundArray(243, 1) = 171
    backgroundArray(243, 2) = 202
    backgroundArray(244, 0) = 140
    backgroundArray(244, 1) = 238
    backgroundArray(244, 2) = 149
    backgroundArray(245, 0) = 104
    backgroundArray(245, 1) = 247
    backgroundArray(245, 2) = 2
    backgroundArray(246, 0) = 147
    backgroundArray(246, 1) = 5
    backgroundArray(246, 2) = 76
    backgroundArray(247, 0) = 175
    backgroundArray(247, 1) = 83
    backgroundArray(247, 2) = 230
    backgroundArray(248, 0) = 191
    backgroundArray(248, 1) = 186
    backgroundArray(248, 2) = 155
    backgroundArray(249, 0) = 221
    backgroundArray(249, 1) = 6
    backgroundArray(249, 2) = 84
    backgroundArray(250, 0) = 197
    backgroundArray(250, 1) = 102
    backgroundArray(250, 2) = 202
    backgroundArray(251, 0) = 134
    backgroundArray(251, 1) = 43
    backgroundArray(251, 2) = 136
    backgroundArray(252, 0) = 202
    backgroundArray(252, 1) = 152
    backgroundArray(252, 2) = 119
    backgroundArray(253, 0) = 186
    backgroundArray(253, 1) = 78
    backgroundArray(253, 2) = 209
    backgroundArray(254, 0) = 23
    backgroundArray(254, 1) = 186
    backgroundArray(254, 2) = 133
    backgroundArray(255, 0) = 165
    backgroundArray(255, 1) = 146
    backgroundArray(255, 2) = 91
    backgroundArray(256, 0) = 32
    backgroundArray(256, 1) = 141
    backgroundArray(256, 2) = 122
    backgroundArray(257, 0) = 186
    backgroundArray(257, 1) = 5
    backgroundArray(257, 2) = 30
    backgroundArray(258, 0) = 192
    backgroundArray(258, 1) = 164
    backgroundArray(258, 2) = 9
    backgroundArray(259, 0) = 215
    backgroundArray(259, 1) = 51
    backgroundArray(259, 2) = 50
    backgroundArray(260, 0) = 199
    backgroundArray(260, 1) = 192
    backgroundArray(260, 2) = 49
    backgroundArray(261, 0) = 232
    backgroundArray(261, 1) = 2
    backgroundArray(261, 2) = 168
    backgroundArray(262, 0) = 171
    backgroundArray(262, 1) = 225
    backgroundArray(262, 2) = 209
    backgroundArray(263, 0) = 133
    backgroundArray(263, 1) = 144
    backgroundArray(263, 2) = 81
    backgroundArray(264, 0) = 252
    backgroundArray(264, 1) = 155
    backgroundArray(264, 2) = 134
    backgroundArray(265, 0) = 116
    backgroundArray(265, 1) = 46
    backgroundArray(265, 2) = 21
    backgroundArray(266, 0) = 77
    backgroundArray(266, 1) = 198
    backgroundArray(266, 2) = 16
    backgroundArray(267, 0) = 13
    backgroundArray(267, 1) = 98
    backgroundArray(267, 2) = 60
    backgroundArray(268, 0) = 161
    backgroundArray(268, 1) = 156
    backgroundArray(268, 2) = 62
    backgroundArray(269, 0) = 250
    backgroundArray(269, 1) = 5
    backgroundArray(269, 2) = 253
    backgroundArray(270, 0) = 73
    backgroundArray(270, 1) = 3
    backgroundArray(270, 2) = 98
    backgroundArray(271, 0) = 53
    backgroundArray(271, 1) = 159
    backgroundArray(271, 2) = 233
    backgroundArray(272, 0) = 17
    backgroundArray(272, 1) = 30
    backgroundArray(272, 2) = 201
    backgroundArray(273, 0) = 133
    backgroundArray(273, 1) = 27
    backgroundArray(273, 2) = 158
    backgroundArray(274, 0) = 187
    backgroundArray(274, 1) = 3
    backgroundArray(274, 2) = 87
    backgroundArray(275, 0) = 55
    backgroundArray(275, 1) = 122
    backgroundArray(275, 2) = 183
    backgroundArray(276, 0) = 211
    backgroundArray(276, 1) = 228
    backgroundArray(276, 2) = 8
    backgroundArray(277, 0) = 244
    backgroundArray(277, 1) = 181
    backgroundArray(277, 2) = 254
    backgroundArray(278, 0) = 218
    backgroundArray(278, 1) = 84
    backgroundArray(278, 2) = 2
    backgroundArray(279, 0) = 214
    backgroundArray(279, 1) = 14
    backgroundArray(279, 2) = 57
    backgroundArray(280, 0) = 216
    backgroundArray(280, 1) = 218
    backgroundArray(280, 2) = 253
    backgroundArray(281, 0) = 153
    backgroundArray(281, 1) = 92
    backgroundArray(281, 2) = 115
    backgroundArray(282, 0) = 28
    backgroundArray(282, 1) = 69
    backgroundArray(282, 2) = 133
    backgroundArray(283, 0) = 191
    backgroundArray(283, 1) = 207
    backgroundArray(283, 2) = 196
    backgroundArray(284, 0) = 33
    backgroundArray(284, 1) = 251
    backgroundArray(284, 2) = 111
    backgroundArray(285, 0) = 158
    backgroundArray(285, 1) = 147
    backgroundArray(285, 2) = 55
    backgroundArray(286, 0) = 146
    backgroundArray(286, 1) = 232
    backgroundArray(286, 2) = 122
    backgroundArray(287, 0) = 103
    backgroundArray(287, 1) = 219
    backgroundArray(287, 2) = 67
    backgroundArray(288, 0) = 61
    backgroundArray(288, 1) = 84
    backgroundArray(288, 2) = 180
    backgroundArray(289, 0) = 248
    backgroundArray(289, 1) = 3
    backgroundArray(289, 2) = 110
    backgroundArray(290, 0) = 178
    backgroundArray(290, 1) = 73
    backgroundArray(290, 2) = 134
    backgroundArray(291, 0) = 117
    backgroundArray(291, 1) = 161
    backgroundArray(291, 2) = 125
    backgroundArray(292, 0) = 28
    backgroundArray(292, 1) = 99
    backgroundArray(292, 2) = 159
    backgroundArray(293, 0) = 12
    backgroundArray(293, 1) = 225
    backgroundArray(293, 2) = 139
    backgroundArray(294, 0) = 38
    backgroundArray(294, 1) = 50
    backgroundArray(294, 2) = 199
    backgroundArray(295, 0) = 195
    backgroundArray(295, 1) = 216
    backgroundArray(295, 2) = 106
    backgroundArray(296, 0) = 179
    backgroundArray(296, 1) = 80
    backgroundArray(296, 2) = 192
    backgroundArray(297, 0) = 120
    backgroundArray(297, 1) = 146
    backgroundArray(297, 2) = 24
    backgroundArray(298, 0) = 53
    backgroundArray(298, 1) = 39
    backgroundArray(298, 2) = 241
    backgroundArray(299, 0) = 50
    backgroundArray(299, 1) = 210
    backgroundArray(299, 2) = 39
    backgroundArray(300, 0) = 213
    backgroundArray(300, 1) = 246
    backgroundArray(300, 2) = 254
    backgroundArray(301, 0) = 35
    backgroundArray(301, 1) = 20
    backgroundArray(301, 2) = 70
    backgroundArray(302, 0) = 207
    backgroundArray(302, 1) = 1
    backgroundArray(302, 2) = 9
    backgroundArray(303, 0) = 174
    backgroundArray(303, 1) = 204
    backgroundArray(303, 2) = 64
    backgroundArray(304, 0) = 144
    backgroundArray(304, 1) = 76
    backgroundArray(304, 2) = 36
    backgroundArray(305, 0) = 184
    backgroundArray(305, 1) = 163
    backgroundArray(305, 2) = 173
    backgroundArray(306, 0) = 142
    backgroundArray(306, 1) = 239
    backgroundArray(306, 2) = 165
    backgroundArray(307, 0) = 162
    backgroundArray(307, 1) = 230
    backgroundArray(307, 2) = 216
    backgroundArray(308, 0) = 233
    backgroundArray(308, 1) = 121
    backgroundArray(308, 2) = 106
    backgroundArray(309, 0) = 163
    backgroundArray(309, 1) = 124
    backgroundArray(309, 2) = 255
    backgroundArray(310, 0) = 18
    backgroundArray(310, 1) = 142
    backgroundArray(310, 2) = 194
    backgroundArray(311, 0) = 239
    backgroundArray(311, 1) = 177
    backgroundArray(311, 2) = 78
    backgroundArray(312, 0) = 225
    backgroundArray(312, 1) = 44
    backgroundArray(312, 2) = 25
    backgroundArray(313, 0) = 198
    backgroundArray(313, 1) = 17
    backgroundArray(313, 2) = 151
    backgroundArray(314, 0) = 80
    backgroundArray(314, 1) = 155
    backgroundArray(314, 2) = 36
    backgroundArray(315, 0) = 40
    backgroundArray(315, 1) = 103
    backgroundArray(315, 2) = 71
    backgroundArray(316, 0) = 247
    backgroundArray(316, 1) = 203
    backgroundArray(316, 2) = 75
    backgroundArray(317, 0) = 34
    backgroundArray(317, 1) = 220
    backgroundArray(317, 2) = 112
    backgroundArray(318, 0) = 92
    backgroundArray(318, 1) = 58
    backgroundArray(318, 2) = 172
    backgroundArray(319, 0) = 187
    backgroundArray(319, 1) = 253
    backgroundArray(319, 2) = 63
    backgroundArray(320, 0) = 63
    backgroundArray(320, 1) = 160
    backgroundArray(320, 2) = 96
    backgroundArray(321, 0) = 165
    backgroundArray(321, 1) = 70
    backgroundArray(321, 2) = 196
    backgroundArray(322, 0) = 178
    backgroundArray(322, 1) = 142
    backgroundArray(322, 2) = 161
    backgroundArray(323, 0) = 150
    backgroundArray(323, 1) = 82
    backgroundArray(323, 2) = 167
    backgroundArray(324, 0) = 136
    backgroundArray(324, 1) = 182
    backgroundArray(324, 2) = 44
    backgroundArray(325, 0) = 148
    backgroundArray(325, 1) = 2
    backgroundArray(325, 2) = 1
    backgroundArray(326, 0) = 190
    backgroundArray(326, 1) = 111
    backgroundArray(326, 2) = 183
    backgroundArray(327, 0) = 161
    backgroundArray(327, 1) = 219
    backgroundArray(327, 2) = 27
    backgroundArray(328, 0) = 97
    backgroundArray(328, 1) = 142
    backgroundArray(328, 2) = 128
    backgroundArray(329, 0) = 44
    backgroundArray(329, 1) = 184
    backgroundArray(329, 2) = 254
    backgroundArray(330, 0) = 250
    backgroundArray(330, 1) = 73
    backgroundArray(330, 2) = 188
    backgroundArray(331, 0) = 165
    backgroundArray(331, 1) = 180
    backgroundArray(331, 2) = 201
    backgroundArray(332, 0) = 150
    backgroundArray(332, 1) = 60
    backgroundArray(332, 2) = 242
    backgroundArray(333, 0) = 197
    backgroundArray(333, 1) = 25
    backgroundArray(333, 2) = 188
    backgroundArray(334, 0) = 236
    backgroundArray(334, 1) = 4
    backgroundArray(334, 2) = 235
    backgroundArray(335, 0) = 241
    backgroundArray(335, 1) = 89
    backgroundArray(335, 2) = 239
    backgroundArray(336, 0) = 117
    backgroundArray(336, 1) = 77
    backgroundArray(336, 2) = 147
    backgroundArray(337, 0) = 93
    backgroundArray(337, 1) = 206
    backgroundArray(337, 2) = 106
    backgroundArray(338, 0) = 217
    backgroundArray(338, 1) = 52
    backgroundArray(338, 2) = 218
    backgroundArray(339, 0) = 77
    backgroundArray(339, 1) = 175
    backgroundArray(339, 2) = 221
    backgroundArray(340, 0) = 135
    backgroundArray(340, 1) = 39
    backgroundArray(340, 2) = 226
    backgroundArray(341, 0) = 61
    backgroundArray(341, 1) = 202
    backgroundArray(341, 2) = 101
    backgroundArray(342, 0) = 148
    backgroundArray(342, 1) = 55
    backgroundArray(342, 2) = 168
    backgroundArray(343, 0) = 20
    backgroundArray(343, 1) = 98
    backgroundArray(343, 2) = 120
    backgroundArray(344, 0) = 171
    backgroundArray(344, 1) = 159
    backgroundArray(344, 2) = 58
    backgroundArray(345, 0) = 128
    backgroundArray(345, 1) = 29
    backgroundArray(345, 2) = 31
    backgroundArray(346, 0) = 54
    backgroundArray(346, 1) = 65
    backgroundArray(346, 2) = 138
    backgroundArray(347, 0) = 16
    backgroundArray(347, 1) = 31
    backgroundArray(347, 2) = 52
    backgroundArray(348, 0) = 177
    backgroundArray(348, 1) = 167
    backgroundArray(348, 2) = 37
    backgroundArray(349, 0) = 90
    backgroundArray(349, 1) = 201
    backgroundArray(349, 2) = 133
    backgroundArray(350, 0) = 168
    backgroundArray(350, 1) = 32
    backgroundArray(350, 2) = 141
    backgroundArray(351, 0) = 30
    backgroundArray(351, 1) = 54
    backgroundArray(351, 2) = 47
    backgroundArray(352, 0) = 4
    backgroundArray(352, 1) = 19
    backgroundArray(352, 2) = 3
    backgroundArray(353, 0) = 132
    backgroundArray(353, 1) = 86
    backgroundArray(353, 2) = 201
    backgroundArray(354, 0) = 105
    backgroundArray(354, 1) = 153
    backgroundArray(354, 2) = 47
    backgroundArray(355, 0) = 19
    backgroundArray(355, 1) = 139
    backgroundArray(355, 2) = 233
    backgroundArray(356, 0) = 66
    backgroundArray(356, 1) = 24
    backgroundArray(356, 2) = 124
    backgroundArray(357, 0) = 103
    backgroundArray(357, 1) = 100
    backgroundArray(357, 2) = 148
    backgroundArray(358, 0) = 30
    backgroundArray(358, 1) = 115
    backgroundArray(358, 2) = 177
    backgroundArray(359, 0) = 61
    backgroundArray(359, 1) = 237
    backgroundArray(359, 2) = 83
    backgroundArray(360, 0) = 208
    backgroundArray(360, 1) = 96
    backgroundArray(360, 2) = 218
    backgroundArray(361, 0) = 68
    backgroundArray(361, 1) = 90
    backgroundArray(361, 2) = 64
    backgroundArray(362, 0) = 131
    backgroundArray(362, 1) = 12
    backgroundArray(362, 2) = 103
    backgroundArray(363, 0) = 6
    backgroundArray(363, 1) = 135
    backgroundArray(363, 2) = 25
    backgroundArray(364, 0) = 191
    backgroundArray(364, 1) = 52
    backgroundArray(364, 2) = 0
    backgroundArray(365, 0) = 43
    backgroundArray(365, 1) = 110
    backgroundArray(365, 2) = 200
    backgroundArray(366, 0) = 178
    backgroundArray(366, 1) = 135
    backgroundArray(366, 2) = 105
    backgroundArray(367, 0) = 144
    backgroundArray(367, 1) = 19
    backgroundArray(367, 2) = 177
    backgroundArray(368, 0) = 72
    backgroundArray(368, 1) = 144
    backgroundArray(368, 2) = 210
    backgroundArray(369, 0) = 6
    backgroundArray(369, 1) = 3
    backgroundArray(369, 2) = 82
    backgroundArray(370, 0) = 184
    backgroundArray(370, 1) = 129
    backgroundArray(370, 2) = 122
    backgroundArray(371, 0) = 142
    backgroundArray(371, 1) = 219
    backgroundArray(371, 2) = 83
    backgroundArray(372, 0) = 3
    backgroundArray(372, 1) = 120
    backgroundArray(372, 2) = 196
    backgroundArray(373, 0) = 5
    backgroundArray(373, 1) = 184
    backgroundArray(373, 2) = 136
    backgroundArray(374, 0) = 16
    backgroundArray(374, 1) = 233
    backgroundArray(374, 2) = 109
    backgroundArray(375, 0) = 21
    backgroundArray(375, 1) = 83
    backgroundArray(375, 2) = 83
    backgroundArray(376, 0) = 106
    backgroundArray(376, 1) = 212
    backgroundArray(376, 2) = 161
    backgroundArray(377, 0) = 46
    backgroundArray(377, 1) = 216
    backgroundArray(377, 2) = 123
    backgroundArray(378, 0) = 195
    backgroundArray(378, 1) = 49
    backgroundArray(378, 2) = 48
    backgroundArray(379, 0) = 138
    backgroundArray(379, 1) = 250
    backgroundArray(379, 2) = 0
    backgroundArray(380, 0) = 211
    backgroundArray(380, 1) = 125
    backgroundArray(380, 2) = 44
    backgroundArray(381, 0) = 27
    backgroundArray(381, 1) = 105
    backgroundArray(381, 2) = 218
    backgroundArray(382, 0) = 119
    backgroundArray(382, 1) = 81
    backgroundArray(382, 2) = 205
    backgroundArray(383, 0) = 33
    backgroundArray(383, 1) = 219
    backgroundArray(383, 2) = 226
    backgroundArray(384, 0) = 33
    backgroundArray(384, 1) = 195
    backgroundArray(384, 2) = 226
    backgroundArray(385, 0) = 54
    backgroundArray(385, 1) = 157
    backgroundArray(385, 2) = 25
    backgroundArray(386, 0) = 24
    backgroundArray(386, 1) = 106
    backgroundArray(386, 2) = 121
    backgroundArray(387, 0) = 15
    backgroundArray(387, 1) = 105
    backgroundArray(387, 2) = 223
    backgroundArray(388, 0) = 107
    backgroundArray(388, 1) = 191
    backgroundArray(388, 2) = 24
    backgroundArray(389, 0) = 171
    backgroundArray(389, 1) = 231
    backgroundArray(389, 2) = 221
    backgroundArray(390, 0) = 136
    backgroundArray(390, 1) = 61
    backgroundArray(390, 2) = 192
    backgroundArray(391, 0) = 124
    backgroundArray(391, 1) = 26
    backgroundArray(391, 2) = 76
    backgroundArray(392, 0) = 73
    backgroundArray(392, 1) = 63
    backgroundArray(392, 2) = 118
    backgroundArray(393, 0) = 164
    backgroundArray(393, 1) = 236
    backgroundArray(393, 2) = 201
    backgroundArray(394, 0) = 136
    backgroundArray(394, 1) = 59
    backgroundArray(394, 2) = 181
    backgroundArray(395, 0) = 72
    backgroundArray(395, 1) = 63
    backgroundArray(395, 2) = 31
    backgroundArray(396, 0) = 15
    backgroundArray(396, 1) = 136
    backgroundArray(396, 2) = 87
    backgroundArray(397, 0) = 100
    backgroundArray(397, 1) = 128
    backgroundArray(397, 2) = 177
    backgroundArray(398, 0) = 179
    backgroundArray(398, 1) = 189
    backgroundArray(398, 2) = 242
    backgroundArray(399, 0) = 59
    backgroundArray(399, 1) = 108
    backgroundArray(399, 2) = 37

    ' Randomly pick one of the 400 background colors to return
    backgroundColorIndex = Int((400) * Rnd)

    ' Assign the RGB values of the background color to the array
    returnedBackgroundArray(0) = backgroundArray(backgroundColorIndex, 0)
    returnedBackgroundArray(1) = backgroundArray(backgroundColorIndex, 1)
    returnedBackgroundArray(2) = backgroundArray(backgroundColorIndex, 2)

    ' Return the background color array to the main program
    GenerateBackgroundColor = returnedBackgroundArray

End Function

Function CalcLastRow(Optional sheet As Variant, Optional col As Variant) As Double

' Calculate the last row used.
'
' CalcLastRow("Sheet1", 3)  will calculate the last row used in column 3 in sheet named "Sheet1"
' CalcLastRow("Sheet1")     will calculate the last row used in sheet named "Sheet1"
' CalcLastRow(3)            will calculate the last row used in column 3 in ActiveSheet
' CalcLastRow()             will calculate the last row used in ActiveSheet

Dim lastCellData As Range

' If sheet and col are not input then calculate the last row used in ActiveSheet.
If IsMissing(sheet) And IsMissing(col) Then

    Set lastCellData = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    CalcLastRow = lastCellData.row

' If both the sheet and col are input then calculate the last row in the specific
' col column and sheet worksheet.
ElseIf TypeName(sheet) = "String" And TypeName(col) = "Integer" Then
    
    CalcLastRow = Sheets(sheet).Cells(Rows.Count, col).End(xlUp).row

' If no sheet is input but the col is input then calculate the last row in the
' column col in the ActiveSheet.
ElseIf TypeName(sheet) = "Integer" Then
    
    col = sheet
    CalcLastRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row

' If no col is input but the sheet is input then calculate the last row in the
' entire worksheet for the col in the sheet worksheet.
ElseIf TypeName(sheet) = "String" Then

    Set lastCellData = Sheets(sheet).Cells.Find(What:="*", After:=Sheets(sheet).Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    CalcLastRow = lastCellData.row

End If

End Function

Function CalcLastColumn(Optional sheet As Variant, Optional row As Variant) As Double

' Calculate the last column used.
'
' CalcLastColumn("Sheet1", 3)will calculate the last column used in row 3 in sheet named "Sheet1"
' CalcLastColumn("Sheet1")   will calculate the last column used in sheet named "Sheet1"
' CalcLastColumn(3)   will calculate the last column used in row 3 in ActiveSheet
' CalcLastColumn()   will calculate the last column used in ActiveSheet

Dim lastCellData As Range

' If sheet and row are not input then calculate the last column used in ActiveSheet.
If IsMissing(sheet) And IsMissing(row) Then

    Set lastCellData = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    CalcLastColumn = lastCellData.Column

' If both the sheet and row are input then calculate the last column in the specific
' row row and sheet worksheet.
ElseIf TypeName(sheet) = "String" And TypeName(row) = "Integer" Then
    
    CalcLastColumn = Sheets(sheet).Cells(row, Columns.Count).End(xlToLeft).Column

' If no sheet is input but the row is input then calculate the last row in the
' row row in the ActiveSheet.
ElseIf TypeName(sheet) = "Integer" Then
    
    row = sheet
    CalcLastColumn = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column

' If no row is input but the sheet is input then calculate the last column in the
' entire worksheet for the row in the sheet worksheet.
ElseIf TypeName(sheet) = "String" Then

    Set lastCellData = Sheets(sheet).Cells.Find(What:="*", After:=Sheets(sheet).Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    CalcLastColumn = lastCellData.Column

End If

End Function



