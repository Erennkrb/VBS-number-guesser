' === The Ultimate Number Guessing Game v2 ===
Option Explicit

Dim fso, scoreFile, bestScores
Dim totalWins, totalLosses, streak, bestStreak, totalPoints
Dim playAgain

Set fso = CreateObject("Scripting.FileSystemObject")
Set bestScores = CreateObject("Scripting.Dictionary")
scoreFile = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "guess_highscores.txt")

totalWins = 0
totalLosses = 0
streak = 0
bestStreak = 0
totalPoints = 0

LoadScores

MsgBox "Welcome to the Number Guessing Game!" & vbCrLf & vbCrLf & _
       "Pick a difficulty, then find the secret number." & vbCrLf & _
       "Hints tell you how hot or cold your guess is." & vbCrLf & _
       "Press Cancel in the guess box to quit a round.", vbInformation, "Number Guessing Game"

Do
    PlayRound
    playAgain = MsgBox("Play another round?" & vbCrLf & vbCrLf & StatsText(), vbYesNo + vbQuestion, "Play Again?")
Loop While playAgain = vbYes

MsgBox "Thanks for playing!" & vbCrLf & vbCrLf & StatsText(), vbInformation, "Goodbye"
WScript.Quit

' ---------------------------------------------------------------

Sub PlayRound()
    Dim levelName, lo, hi, maxAttempts, multiplier
    Dim target, attempts, won, quitRound
    Dim raw, guess, valid, errMsg, prompt
    Dim history, diff, span, hint, points, key

    If Not ChooseDifficulty(levelName, lo, hi, maxAttempts, multiplier) Then Exit Sub

    Randomize
    target = Int((hi - lo + 1) * Rnd) + lo
    span = hi - lo + 1
    attempts = 0
    won = False
    quitRound = False
    history = ""
    hint = ""

    Do While attempts < maxAttempts And Not won And Not quitRound
        valid = False
        errMsg = ""

        Do Until valid
            prompt = "Difficulty: " & levelName & "  (" & lo & " to " & hi & ")" & vbCrLf & _
                     "Attempt " & (attempts + 1) & " of " & maxAttempts & vbCrLf
            If hint <> "" Then prompt = prompt & vbCrLf & hint & vbCrLf
            If history <> "" Then prompt = prompt & "Your guesses: " & history & vbCrLf
            If errMsg <> "" Then prompt = prompt & vbCrLf & "! " & errMsg & vbCrLf
            prompt = prompt & vbCrLf & "Enter your guess:"

            raw = Trim(InputBox(prompt, "Guess the Number"))

            If raw = "" Then
                If MsgBox("Give up this round?", vbYesNo + vbQuestion, "Quit round") = vbYes Then
                    quitRound = True
                    valid = True
                End If
            ElseIf Not IsNumeric(raw) Then
                errMsg = "That is not a number."
            ElseIf CDbl(raw) <> Int(CDbl(raw)) Then
                errMsg = "Whole numbers only."
            Else
                guess = CLng(raw)
                If guess < lo Or guess > hi Then
                    errMsg = "Stay between " & lo & " and " & hi & "."
                ElseIf InStr("," & Replace(history, " ", "") & ",", "," & guess & ",") > 0 Then
                    errMsg = "You already tried " & guess & "."
                Else
                    valid = True
                End If
            End If
        Loop

        If Not quitRound Then
            attempts = attempts + 1
            If history <> "" Then history = history & ", "
            history = history & guess

            If guess = target Then
                won = True
            Else
                diff = Abs(guess - target)
                If diff <= span * 0.03 Then
                    hint = "BURNING HOT"
                ElseIf diff <= span * 0.1 Then
                    hint = "Hot"
                ElseIf diff <= span * 0.25 Then
                    hint = "Warm"
                Else
                    hint = "Cold"
                End If
                If guess < target Then
                    hint = hint & " - go HIGHER"
                Else
                    hint = hint & " - go LOWER"
                End If
            End If
        End If
    Loop

    If won Then
        points = (maxAttempts - attempts + 1) * 100 * multiplier
        totalPoints = totalPoints + points
        totalWins = totalWins + 1
        streak = streak + 1
        If streak > bestStreak Then bestStreak = streak

        key = levelName
        Dim msg, isRecord
        isRecord = False
        If Not bestScores.Exists(key) Then
            isRecord = True
        ElseIf attempts < bestScores(key) Then
            isRecord = True
        End If
        If isRecord Then
            bestScores(key) = attempts
            SaveScores
        End If

        msg = "YOU WIN!" & vbCrLf & vbCrLf & _
              "The number was " & target & "." & vbCrLf & _
              "Solved in " & attempts & " attempt(s)." & vbCrLf & _
              "Points earned: " & points
        If isRecord Then msg = msg & vbCrLf & vbCrLf & "NEW PERSONAL BEST on " & levelName & "!"
        MsgBox msg, vbInformation, "Victory"
    ElseIf quitRound Then
        streak = 0
        MsgBox "Round abandoned." & vbCrLf & "The number was " & target & ".", vbExclamation, "Round Over"
    Else
        totalLosses = totalLosses + 1
        streak = 0
        MsgBox "Game Over, out of attempts." & vbCrLf & vbCrLf & _
               "The secret number was " & target & ".", vbCritical, "Out of Turns"
    End If
End Sub

Function ChooseDifficulty(ByRef levelName, ByRef lo, ByRef hi, ByRef maxAttempts, ByRef multiplier)
    Dim choice, menu
    ChooseDifficulty = False

    Do
        menu = "Choose a difficulty:" & vbCrLf & vbCrLf & _
               "1  Easy     (1 to 50, 10 attempts)" & BestText("Easy") & vbCrLf & _
               "2  Normal   (1 to 100, 7 attempts)" & BestText("Normal") & vbCrLf & _
               "3  Hard     (1 to 500, 9 attempts)" & BestText("Hard") & vbCrLf & _
               "4  Insane   (1 to 1000, 10 attempts)" & BestText("Insane") & vbCrLf & vbCrLf & _
               "Enter 1, 2, 3 or 4:"

        choice = Trim(InputBox(menu, "Difficulty", "2"))
        If choice = "" Then Exit Function

        Select Case choice
            Case "1": levelName = "Easy":   lo = 1: hi = 50:   maxAttempts = 10: multiplier = 1: ChooseDifficulty = True
            Case "2": levelName = "Normal": lo = 1: hi = 100:  maxAttempts = 7:  multiplier = 2: ChooseDifficulty = True
            Case "3": levelName = "Hard":   lo = 1: hi = 500:  maxAttempts = 9:  multiplier = 3: ChooseDifficulty = True
            Case "4": levelName = "Insane": lo = 1: hi = 1000: maxAttempts = 10: multiplier = 5: ChooseDifficulty = True
            Case Else: MsgBox "Please enter 1, 2, 3 or 4.", vbExclamation, "Invalid choice"
        End Select
    Loop Until ChooseDifficulty
End Function

Function BestText(levelName)
    If bestScores.Exists(levelName) Then
        BestText = "   [best: " & bestScores(levelName) & "]"
    Else
        BestText = ""
    End If
End Function

Function StatsText()
    StatsText = "Wins: " & totalWins & "   Losses: " & totalLosses & vbCrLf & _
                "Current streak: " & streak & "   Best streak: " & bestStreak & vbCrLf & _
                "Total points: " & totalPoints
End Function

Sub LoadScores()
    Dim f, line, parts
    If Not fso.FileExists(scoreFile) Then Exit Sub
    On Error Resume Next
    Set f = fso.OpenTextFile(scoreFile, 1)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    Do While Not f.AtEndOfStream
        line = f.ReadLine
        parts = Split(line, "=")
        If UBound(parts) = 1 Then
            If IsNumeric(parts(1)) Then bestScores(parts(0)) = CLng(parts(1))
        End If
    Loop
    f.Close
    On Error GoTo 0
End Sub

Sub SaveScores()
    Dim f, k
    On Error Resume Next
    Set f = fso.CreateTextFile(scoreFile, True)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For Each k In bestScores.Keys
        f.WriteLine k & "=" & bestScores(k)
    Next
    f.Close
    On Error GoTo 0
End Sub