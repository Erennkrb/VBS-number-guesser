' --- The Ultimate Number Guessing Game ---
Option Explicit

Dim targetNum, userGuess, attempts, maxAttempts, isGuessed, playAgain

Do
    ' Initialize game variables
    Randomize
    targetNum = Int((100 * Rnd) + 1) ' Random number between 1 and 100
    attempts = 0
    maxAttempts = 7
    isGuessed = False

    MsgBox "Welcome to the Number Guessing Game!" & vbCrLf & vbCrLf & _
           "I have chosen a secret number between 1 and 100." & vbCrLf & _
           "You have " & maxAttempts & " attempts to guess it. Good luck!", vbInformation, "Game Start"

    Do While attempts < maxAttempts And Not isGuessed
        attempts = attempts + 1
        
        userGuess = InputBox("Attempt " & attempts & " of " & maxAttempts & ":" & vbCrLf & vbCrLf & _
                             "Enter your guess (1 - 100):", "Guess the Number")
        
        ' If user clicks Cancel or leaves it empty, exit the game
        If userGuess = "" Then 
            MsgBox "Thanks for playing! Goodbye!", vbInformation, "Game Over"
            WScript.Quit
        End If
        
        ' Validate integer input
        If IsNumeric(userGuess) Then
            userGuess = CInt(userGuess)
            
            If userGuess = targetNum Then
                isGuessed = True
                MsgBox "🎉 Congratulations! You guessed the correct number (" & targetNum & ") in " & attempts & " attempt(s)!", vbInformation, "YOU WIN!"
            ElseIf userGuess < targetNum Then
                MsgBox "Too LOW! Try a higher number. 📈" & vbCrLf & _
                       "Attempts remaining: " & (maxAttempts - attempts), vbExclamation, "Hint"
            Else
                MsgBox "Too HIGH! Try a lower number. 📉" & vbCrLf & _
                       "Attempts remaining: " & (maxAttempts - attempts), vbExclamation, "Hint"
            End If
        Else
            MsgBox "Please enter a valid number!", vbCritical, "Error"
            attempts = attempts - 1 ' Don't penalize for invalid input
        End If
    Loop

    If Not isGuessed Then
        MsgBox "💥 Game Over! You ran out of attempts." & vbCrLf & _
               "The secret number was " & targetNum & ".", vbCritical, "Out of Turns"
    End If

    ' Ask to play again
    playAgain = MsgBox("Do you want to play another round?", vbYesNo + vbQuestion, "Play Again?")

Loop While playAgain = vbYes

MsgBox "Thanks for playing! See you next time.", vbInformation, "Goodbye"
