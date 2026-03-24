Public Class _4InARow
    Dim empty As Image = Image.FromFile("empty.png")
    Dim yellowPiece As Image = Image.FromFile("yellow_piece.png")
    Dim redPiece As Image = Image.FromFile("red_piece.png")
    Dim n As Integer = 6, m As Integer = 7
    Dim cells(n, m) As Button
    Dim board(n, m), lastRowIdx(m) As Integer
    Dim over As Boolean

    Private Sub _4InARow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For j = 0 To m - 1
            For i = 0 To n - 1
                cells(i, j) = DirectCast(Controls.Find("cell" & i & j, False).First, Button)
                cells(i, j).Image = empty
                board(i, j) = 0
            Next
            lastRowIdx(j) = n - 1
        Next
        over = False
    End Sub

    Private Function IsWin(n As Integer, m As Integer, board(,) As Integer, piece As Integer)
        Dim dy() As Integer = {-1, -1, 0, 1, 1, 1, 0, -1}
        Dim dx() As Integer = {0, 1, 1, 1, 0, -1, -1, -1}
        For i = 0 To n - 1
            For j = 0 To m - 1
                If board(i, j) = piece Then
                    For k = 0 To 7
                        If i + dy(k) * 3 >= 0 And i + dy(k) * 3 <= n - 1 And j + dx(k) * 3 >= 0 And j + dx(k) * 3 <= m - 1 Then
                            Dim isPossibleWin As Boolean = True
                            For t = 1 To 3
                                Dim r As Integer = i + dy(k) * t, c As Integer = j + dx(k) * t
                                If board(r, c) <> piece Then
                                    isPossibleWin = False
                                    Exit For
                                End If
                            Next
                            If isPossibleWin Then
                                Return True
                            End If
                        End If
                    Next
                End If
            Next
        Next
        Return False
    End Function

    Private Function IsDraw(m As Integer, lastRowIdx() As Integer)
        For i = 0 To m - 1
            If lastRowIdx(i) <> -1 Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Function SBE(n As Integer, m As Integer, board(,) As Integer, lastRowIdx() As Integer)
        If IsWin(n, m, board, 1) Then
            Return 1000
        ElseIf IsWin(n, m, board, -1) Then
            Return -1000
        ElseIf IsDraw(m, lastRowIdx) Then
            Return 0
        Else
            Dim score As Integer = 0
            Dim dy() As Integer = {0, -1, -1, 0, 1, 1, 1, 0, -1}
            Dim dx() As Integer = {0, 0, 1, 1, 1, 0, -1, -1, -1}
            For i = 0 To n - 1
                For j = 0 To m - 1
                    For k = 0 To 8
                        If i + dy(k) * 3 >= 0 And i + dy(k) * 3 <= n - 1 And j + dx(k) * 3 >= 0 And j + dx(k) * 3 <= m - 1 Then
                            Dim isPossibleWin As Boolean = True
                            Dim hasReds As Boolean = False, hasYellows As Boolean = False
                            For t = 0 To 3
                                Dim r As Integer = i + dy(k) * t, c As Integer = j + dx(k) * t
                                If board(r, c) = 1 Then
                                    hasReds = True
                                ElseIf board(r, c) = -1 Then
                                    hasYellows = True
                                ElseIf r <> lastRowIdx(c) Then
                                    If dx(k) <> 0 Then
                                        isPossibleWin = False
                                        Exit For
                                    End If
                                End If
                            Next
                            If isPossibleWin Then
                                If hasReds Then
                                    score += 1
                                End If
                                If hasYellows Then
                                    score -= 1
                                End If
                            End If
                        End If
                    Next
                Next
            Next

            ''Rows
            'For i = 0 To n - 1
            '    For j = 0 To m - 4
            '        Dim isPossibleWin As Boolean = True
            '        Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '        For k = 0 To 3
            '            If board(i, j + k) = 1 Then
            '                hasReds = True
            '            ElseIf board(i, j + k) = -1 Then
            '                hasYellows = True
            '            ElseIf i <> lastRowIdx(j + k) Then
            '                isPossibleWin = False
            '                Exit For
            '            End If
            '        Next
            '        If isPossibleWin Then
            '            If hasReds Then
            '                score += 1
            '            End If
            '            If hasYellows Then
            '                score -= 1
            '            End If
            '        End If
            '    Next
            'Next

            ''Columns
            'For j = 0 To m - 1
            '    For i = 0 To n - 4
            '        Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '        For k = 0 To 3
            '            If board(i + k, j) = 1 Then
            '                hasReds = True
            '            ElseIf board(i + k, j) = -1 Then
            '                hasYellows = True
            '            End If
            '        Next
            '        If hasReds Then
            '            score += 1
            '        End If
            '        If hasYellows Then
            '            score -= 1
            '        End If
            '    Next
            'Next

            ''Diagonals 1
            'Dim d1Row As Integer = 0, d1Col As Integer = 0
            'Dim d1R As Integer = d1Row, d1C As Integer = d1Col
            'Do While d1R < n - 3 And d1C < m - 3
            '    Dim isPossibleWin As Boolean = True
            '    Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '    For k = 0 To 3
            '        If board(d1R + k, d1C + k) = 1 Then
            '            hasReds = True
            '        ElseIf board(d1R + k, d1C + k) = -1 Then
            '            hasYellows = True
            '        ElseIf d1R + k <> lastRowIdx(d1C + k) Then
            '            isPossibleWin = False
            '            Exit For
            '        End If
            '    Next
            '    If isPossibleWin Then
            '        If hasReds Then
            '            score += 1
            '        End If
            '        If hasYellows Then
            '            score -= 1
            '        End If
            '    End If
            '    d1R += 1
            '    d1C += 1
            'Loop
            'd1Row += 1
            'd1Col += 1
            'Do While d1Row < n - 3 Or d1Col < m - 3
            '    If d1Row < n - 3 Then
            '        Dim i As Integer = d1Row, c As Integer = 0
            '        Do While i < n And c < m
            '            Dim isPossibleWin As Boolean = True
            '            Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '            For k = 0 To 3
            '                If board(i + k, c + k) = 1 Then
            '                    hasReds = True
            '                ElseIf board(i + k, c + k) = -1 Then
            '                    hasYellows = True
            '                ElseIf i + k <> lastRowIdx(c + k) Then
            '                    isPossibleWin = False
            '                    Exit For
            '                End If
            '            Next
            '            If isPossibleWin Then
            '                If hasReds Then
            '                    score += 1
            '                End If
            '                If hasYellows Then
            '                    score -= 1
            '                End If
            '            End If
            '            i += 1
            '            c += 1
            '        Loop
            '        d1Row += 1
            '    End If
            '    If d1Col < m - 3 Then
            '        Dim r As Integer = 0, j As Integer = d1Col
            '        Do While r < n And j < m
            '            Dim isPossibleWin As Boolean = True
            '            Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '            For k = 0 To 3
            '                If board(r + k, j + k) = 1 Then
            '                    hasReds = True
            '                ElseIf board(r + k, j + k) = -1 Then
            '                    hasYellows = True
            '                ElseIf r + k <> lastRowIdx(j + k) Then
            '                    isPossibleWin = False
            '                    Exit For
            '                End If
            '            Next
            '            If isPossibleWin Then
            '                If hasReds Then
            '                    score += 1
            '                End If
            '                If hasYellows Then
            '                    score -= 1
            '                End If
            '            End If
            '            r += 1
            '            j += 1
            '        Loop
            '        d1Col += 1
            '    End If
            'Loop

            ''Diagonals 2
            'Dim d2Row As Integer = 0, d2Col As Integer = m - 1
            'Dim d2R As Integer = d2Row, d2C As Integer = d2Col
            'Do While d2R < n - 3 And d2C > 2
            '    Dim isPossibleWin As Boolean = True
            '    Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '    For k = 0 To 3
            '        If board(d2R + k, d2C - k) = 1 Then
            '            hasReds = True
            '        ElseIf board(d2R + k, d2C - k) = -1 Then
            '            hasYellows = True
            '        ElseIf d2R + k <> lastRowIdx(d2C - k) Then
            '            isPossibleWin = False
            '            Exit For
            '        End If
            '    Next
            '    If isPossibleWin Then
            '        If hasReds Then
            '            score += 1
            '        End If
            '        If hasYellows Then
            '            score -= 1
            '        End If
            '    End If
            '    d1R += 1
            '    d2C -= 1
            'Loop
            'd2Row += 1
            'd2Col -= 1
            'Do While d2Row < n - 3 Or d2Col > 2
            '    If d2Row < n - 3 Then
            '        Dim i As Integer = d2Row, c As Integer = m - 1
            '        Do While i < n And c > -1
            '            Dim isPossibleWin As Boolean = True
            '            Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '            For k = 0 To 3
            '                If board(i + k, c - k) = 1 Then
            '                    hasReds = True
            '                ElseIf board(i + k, c - k) = -1 Then
            '                    hasYellows = True
            '                ElseIf i + k <> lastRowIdx(c - k) Then
            '                    isPossibleWin = False
            '                    Exit For
            '                End If
            '            Next
            '            If isPossibleWin Then
            '                If hasReds Then
            '                    score += 1
            '                End If
            '                If hasYellows Then
            '                    score -= 1
            '                End If
            '            End If
            '            i += 1
            '            c -= 1
            '        Loop
            '        d2Row += 1
            '    End If
            '    If d2Col > 2 Then
            '        Dim r As Integer = 0, j As Integer = d2Col
            '        Do While r < n And j > 2
            '            Dim isPossibleWin As Boolean = True
            '            Dim hasReds As Boolean = False, hasYellows As Boolean = False
            '            For k = 0 To 3
            '                If board(r + k, j - k) = 1 Then
            '                    hasReds = True
            '                ElseIf board(r + k, j - k) = -1 Then
            '                    hasYellows = True
            '                ElseIf r + k <> lastRowIdx(j - k) Then
            '                    isPossibleWin = False
            '                    Exit For
            '                End If
            '            Next
            '            If isPossibleWin Then
            '                If hasReds Then
            '                    score += 1
            '                End If
            '                If hasYellows Then
            '                    score -= 1
            '                End If
            '            End If
            '            r += 1
            '            j -= 1
            '        Loop
            '        d2Col -= 1
            '    End If
            'Loop

            Return score
        End If
    End Function

    Private Function Minimax(depth As Integer, n As Integer, m As Integer, board(,) As Integer, lastRowIdx() As Integer, isMax As Boolean, alpha As Integer, beta As Integer)
        If depth = 3 Or IsWin(n, m, board, 1) Or IsWin(n, m, board, -1) Or IsDraw(m, lastRowIdx) Then
            Dim score As Integer = SBE(n, m, board, lastRowIdx)
            Return score - depth
        Else
            If isMax Then
                Dim bestScore As Integer = -10000
                For i = 0 To m - 1
                    If lastRowIdx(i) <> -1 Then
                        board(lastRowIdx(i), i) = 1
                        lastRowIdx(i) -= 1
                        bestScore = Math.Max(bestScore, Minimax(depth + 1, n, m, board, lastRowIdx, Not isMax, alpha, beta))
                        lastRowIdx(i) += 1
                        board(lastRowIdx(i), i) = 0
                        alpha = Math.Max(alpha, bestScore)
                        If alpha >= beta Then
                            Exit For
                        End If
                    End If
                Next
                Return bestScore
            Else
                Dim bestScore As Integer = 10000
                For i = 0 To m - 1
                    If lastRowIdx(i) <> -1 Then
                        board(lastRowIdx(i), i) = -1
                        lastRowIdx(i) -= 1
                        bestScore = Math.Min(bestScore, Minimax(depth + 1, n, m, board, lastRowIdx, Not isMax, alpha, beta))
                        lastRowIdx(i) += 1
                        board(lastRowIdx(i), i) = 0
                        beta = Math.Min(beta, bestScore)
                        If alpha >= beta Then
                            Exit For
                        End If
                    End If
                Next
                Return bestScore
            End If
        End If
    End Function

    Private Function FindBestMove(n As Integer, m As Integer, board(,) As Integer, lastRowIdx() As Integer)
        Dim bestMove As Integer = -1, bestScore As Integer = -100000
        For i = 0 To m - 1
            If lastRowIdx(i) <> -1 Then
                board(lastRowIdx(i), i) = 1
                lastRowIdx(i) -= 1
                Dim score As Integer = Minimax(0, n, m, board, lastRowIdx, False, -1000, 1000)
                lastRowIdx(i) += 1
                board(lastRowIdx(i), i) = 0
                If score > bestScore Then
                    bestScore = score
                    bestMove = i
                End If
            End If
        Next
        Return bestMove
    End Function

    Private Sub OppMove(n As Integer, m As Integer, cells(,) As Button, board(,) As Integer, lastRowIdx() As Integer)
        Dim bestMove As Integer = FindBestMove(n, m, board, lastRowIdx)
        cells(lastRowIdx(bestMove), bestMove).Image = redPiece
        board(lastRowIdx(bestMove), bestMove) = 1
        lastRowIdx(bestMove) -= 1
        If IsWin(n, m, board, 1) Then
            GameOver()
            MessageBox.Show("Opponent wins!")
        ElseIf IsWin(n, m, board, -1) Then
            GameOver()
            MessageBox.Show("You win!")
        ElseIf IsDraw(m, lastRowIdx) Then
            GameOver()
            MessageBox.Show("Draw!")
        End If
    End Sub

    Private Sub GameOver()
        over = True
    End Sub

    Private Sub button_Click(sender As Object, e As EventArgs) Handles cell56.Click, cell55.Click, cell54.Click, cell53.Click, cell52.Click, cell51.Click, cell50.Click, cell46.Click, cell45.Click, cell44.Click, cell43.Click, cell42.Click, cell41.Click, cell40.Click, cell36.Click, cell35.Click, cell34.Click, cell33.Click, cell32.Click, cell31.Click, cell30.Click, cell26.Click, cell25.Click, cell24.Click, cell23.Click, cell22.Click, cell21.Click, cell20.Click, cell16.Click, cell15.Click, cell14.Click, cell13.Click, cell12.Click, cell11.Click, cell10.Click, cell06.Click, cell05.Click, cell04.Click, cell03.Click, cell02.Click, cell01.Click, cell00.Click
        If Not over Then
            Dim cellName As String = CType(sender, Button).Name
            Dim col As Integer = AscW(cellName(cellName.Length - 1)) - AscW("0")
            If lastRowIdx(col) <> -1 Then
                cells(lastRowIdx(col), col).Image = yellowPiece
                board(lastRowIdx(col), col) = -1
                lastRowIdx(col) -= 1
                If IsWin(n, m, board, -1) Then
                    GameOver()
                    MessageBox.Show("You win!")
                ElseIf IsWin(n, m, board, 1) Then
                    GameOver()
                    MessageBox.Show("Opponent wins!")
                ElseIf IsDraw(m, lastRowIdx) Then
                    GameOver()
                    MessageBox.Show("Draw!")
                Else
                    OppMove(n, m, cells, board, lastRowIdx)
                End If
            End If
        End If
    End Sub

    Private Sub restartBtn_Click(sender As Object, e As EventArgs) Handles restartBtn.Click
        For j = 0 To m - 1
            For i = 0 To n - 1
                cells(i, j).Image = empty
                board(i, j) = 0
            Next
            lastRowIdx(j) = n - 1
        Next
        over = False
    End Sub
End Class
