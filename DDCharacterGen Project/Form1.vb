Public Class frmMain
    'Purpose: VB Final Project Spring 2018
    'Dungeons and Dragons Character Generator 
    'Programmer: Mary Williams
    'Date created: 4/13/2018 Last edit: 4/21/2018
    'Edit History: 4/13/2018 - Started on form design, went through 3 different design ideas.
    '              4/15/2018 - Finished up form design. Coded: filling lists on form load. Selecting and displaying race and class through if statements and functions.
    '                        - Attempted to code getting attribute values. Added benefits from race is working. Problem: All 6 attributes are being given the same random number.
    '                        - Fixed Attributes. Added functions to determine character's speed and hitdice. Those values are displayed in appropriate label.
    '                        - Added function to determine background and generate and display personality based off that. Need to finish: typing out all the personality choices.
    '              4/16/2018 - Finished adding all the personality traits, ideals, bonds, and flaws for the different backgrounds.
    '                        - Editted the size of the labels for displaying personality to be able to show multiple lines of text.   
    '              4/21/2018 - Editted the form to remove labels that were not intended to be there.
    '                        - Added starting gold labels and code
    '                        - Added Character Name choice code if user does not input a name
    '                        - Still to do: Professionalize form, make sure its all up to standards, add extra form explaining
    '              4/22/2018 - Cleaned up form, followed window's standards
    '                        - Added code to clear labels when appropriate
    '                        - Added message box to inform user to input Player Name if they try to click btnGenerate without inputing it
    '                        - Added message box when form is loaded informing what the purpose of the application is
    '                        - Added code to select all in a txtbox when it gains focus
    '                        - Added code to only allow letters, backspace, space, and apostrophe to be input into txtPlayerName and txtCharName

    'Class level array variables for filling lists on load and use in selection structures when btnGenerate_click event happens
    Dim strRace As String() = {"-Random-", "Human", "Dwarf", "Elf", "Halfling", "Dragonborn", "Gnome", "Half-Elf", "Half-Orc", "Tiefling"}
    Dim strClass As String() = {"-Random-", "Barbarian", "Bard", "Cleric", "Druid", "Fighter", "Monk", "Paladin", "Ranger", "Rogue", "Sorcerer", "Warlock", "Wizard"}

    'Class level random variable to use in Attribute assigning functions.
    'I tried to put it within its own function, but because 'Random' is time dependent, it gave all 6 attributes the same value.
    Dim randRoll6 As New Random
    'After some trial, using it like this does work a bit better. Not all 6 have the same value. 

    'Function to get random number 1-100, used for percentages
    Private Function GetDiceRoll100()
        '100 sided die/ for percentages
        Dim randRoll100 As New Random
        Return randRoll100.Next(1, 101)
    End Function

    'Seperate function for choosing race and class instead of having it all under btnGenerate_click event
    Private Function ChooseRace()
        'use GetDiceRoll100 function to generate a percent and use that for choosing race
        Dim intRacePercent As Integer
        intRacePercent = GetDiceRoll100()
        Dim intRaceArrayIndex As Integer

        If intRacePercent >= 1 AndAlso intRacePercent <= 15 Then
            'Set race to Human race: strRace:1
            intRaceArrayIndex = 1
        ElseIf intRacePercent >= 16 AndAlso intRacePercent <= 30 Then
            'Set race to Dwarf race: strRace:2
            intRaceArrayIndex = 2
        ElseIf intRacePercent >= 31 AndAlso intRacePercent <= 45 Then
            'Set race to Elf race: strRace:3
            intRaceArrayIndex = 3
        ElseIf intRacePercent >= 46 AndAlso intRacePercent <= 60 Then
            'Set race to Halfling race: strRace:4
            intRaceArrayIndex = 4
        ElseIf intRacePercent >= 61 AndAlso intRacePercent <= 68 Then
            'Set race to Dragonborn race: strRace:5
            intRaceArrayIndex = 5
        ElseIf intRacePercent >= 69 AndAlso intRacePercent <= 76 Then
            'Set race to Gnome race: strRace:6
            intRaceArrayIndex = 6
        ElseIf intRacePercent >= 77 AndAlso intRacePercent <= 84 Then
            'Set race to Half-Elf race: strRace7
            intRaceArrayIndex = 7
        ElseIf intRacePercent >= 85 AndAlso intRacePercent <= 92 Then
            'Set race to Half-Orc race: strRace:8
            intRaceArrayIndex = 8
        ElseIf intRacePercent >= 93 AndAlso intRacePercent <= 100 Then
            'Set race to Tiefling race: strRace:9
            intRaceArrayIndex = 9
        End If
        Return intRaceArrayIndex
    End Function

    Private Function ChooseClass(ByVal intRaceNum As Integer)
        'use GetDiceRoll100 function to generate a percent and use that for choosing class based on what race is already assigned
        Dim intClassPercent As Integer
        intClassPercent = GetDiceRoll100()
        Dim intClassArrayIndex As Integer

        'Outer If selection statement for matching it to race
        If intRaceNum = 1 Then
            'Inner if selection statement for Human race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 8 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 9 AndAlso intClassPercent <= 16 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 17 AndAlso intClassPercent <= 24 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 25 AndAlso intClassPercent <= 32 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 33 AndAlso intClassPercent <= 40 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 41 AndAlso intClassPercent <= 48 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 49 AndAlso intClassPercent <= 56 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 57 AndAlso intClassPercent <= 64 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 65 AndAlso intClassPercent <= 72 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 73 AndAlso intClassPercent <= 81 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 82 AndAlso intClassPercent <= 90 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 91 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If
        ElseIf intRaceNum = 2 Then
            'Inner if selection statement for Dwarf race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 13 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 14 AndAlso intClassPercent <= 20 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 21 AndAlso intClassPercent <= 27 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 28 AndAlso intClassPercent <= 34 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 35 AndAlso intClassPercent <= 46 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 47 AndAlso intClassPercent <= 58 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 59 AndAlso intClassPercent <= 65 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 66 AndAlso intClassPercent <= 72 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 73 AndAlso intClassPercent <= 79 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 80 AndAlso intClassPercent <= 86 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 87 AndAlso intClassPercent <= 93 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 94 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 3 Then
            'Inner if selection statement for Elf race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 7 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 8 AndAlso intClassPercent <= 14 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 15 AndAlso intClassPercent <= 21 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 22 AndAlso intClassPercent <= 28 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 29 AndAlso intClassPercent <= 39 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 40 AndAlso intClassPercent <= 50 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 51 AndAlso intClassPercent <= 57 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 58 AndAlso intClassPercent <= 68 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 69 AndAlso intClassPercent <= 79 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 80 AndAlso intClassPercent <= 86 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 87 AndAlso intClassPercent <= 93 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 94 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 4 Then
            'Inner if selection statement for Halfling race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 7 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 8 AndAlso intClassPercent <= 14 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 15 AndAlso intClassPercent <= 21 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 22 AndAlso intClassPercent <= 28 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 29 AndAlso intClassPercent <= 39 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 40 AndAlso intClassPercent <= 50 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 51 AndAlso intClassPercent <= 57 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 58 AndAlso intClassPercent <= 68 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 69 AndAlso intClassPercent <= 79 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 80 AndAlso intClassPercent <= 86 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 87 AndAlso intClassPercent <= 93 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 94 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 5 Then
            'Inner if selection statement for Dragonborn race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 7 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 8 AndAlso intClassPercent <= 14 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 15 AndAlso intClassPercent <= 21 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 22 AndAlso intClassPercent <= 28 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 29 AndAlso intClassPercent <= 35 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 36 AndAlso intClassPercent <= 42 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 43 AndAlso intClassPercent <= 65 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 66 AndAlso intClassPercent <= 72 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 73 AndAlso intClassPercent <= 79 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 80 AndAlso intClassPercent <= 86 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 87 AndAlso intClassPercent <= 93 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 94 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 6 Then
            'Inner if selection statement for Gnome race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 7 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 8 AndAlso intClassPercent <= 14 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 15 AndAlso intClassPercent <= 21 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 22 AndAlso intClassPercent <= 28 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 29 AndAlso intClassPercent <= 35 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 36 AndAlso intClassPercent <= 42 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 43 AndAlso intClassPercent <= 49 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 50 AndAlso intClassPercent <= 56 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 57 AndAlso intClassPercent <= 63 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 64 AndAlso intClassPercent <= 70 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 71 AndAlso intClassPercent <= 77 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 78 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 7 Then
            'Inner if selection statement for Half-Elf race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 7 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 8 AndAlso intClassPercent <= 19 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 20 AndAlso intClassPercent <= 26 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 27 AndAlso intClassPercent <= 33 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 34 AndAlso intClassPercent <= 40 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 41 AndAlso intClassPercent <= 47 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 48 AndAlso intClassPercent <= 54 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 55 AndAlso intClassPercent <= 61 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 62 AndAlso intClassPercent <= 68 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 69 AndAlso intClassPercent <= 80 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 81 AndAlso intClassPercent <= 92 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 93 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 8 Then
            'Inner if selection statement for Half-Orc race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 12 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 13 AndAlso intClassPercent <= 19 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 20 AndAlso intClassPercent <= 26 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 27 AndAlso intClassPercent <= 33 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 34 AndAlso intClassPercent <= 45 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 46 AndAlso intClassPercent <= 52 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 53 AndAlso intClassPercent <= 65 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 66 AndAlso intClassPercent <= 72 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 73 AndAlso intClassPercent <= 79 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 80 AndAlso intClassPercent <= 86 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 87 AndAlso intClassPercent <= 93 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 94 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        ElseIf intRaceNum = 9 Then
            'Inner if selection statement for Tiefling race's class percentage
            If intClassPercent >= 1 AndAlso intClassPercent <= 7 Then
                'Set intClassArrayIndex variable to 1 for Barbarian
                intClassArrayIndex = 1
            ElseIf intClassPercent >= 8 AndAlso intClassPercent <= 22 Then
                'Set intClassArrayIndex variable to 2 for Bard
                intClassArrayIndex = 2
            ElseIf intClassPercent >= 23 AndAlso intClassPercent <= 29 Then
                'Set intClassArrayIndex variable to 3 for Cleric
                intClassArrayIndex = 3
            ElseIf intClassPercent >= 30 AndAlso intClassPercent <= 36 Then
                'Set intClassArrayIndex variable to 4 for Druid
                intClassArrayIndex = 4
            ElseIf intClassPercent >= 37 AndAlso intClassPercent <= 43 Then
                'Set intClassArrayIndex variable to 5 for Fighter
                intClassArrayIndex = 5
            ElseIf intClassPercent >= 44 AndAlso intClassPercent <= 50 Then
                'Set intClassArrayIndex variable to 6 for Monk
                intClassArrayIndex = 6
            ElseIf intClassPercent >= 51 AndAlso intClassPercent <= 57 Then
                'Set intClassArrayIndex variable to 7 for Paladin
                intClassArrayIndex = 7
            ElseIf intClassPercent >= 58 AndAlso intClassPercent <= 64 Then
                'Set intClassArrayIndex variable to 8 for Ranger
                intClassArrayIndex = 8
            ElseIf intClassPercent >= 65 AndAlso intClassPercent <= 71 Then
                'Set intClassArrayIndex variable to 9 for Rogue
                intClassArrayIndex = 9
            ElseIf intClassPercent >= 72 AndAlso intClassPercent <= 78 Then
                'Set intClassArrayIndex variable to 10 for Sorcerer
                intClassArrayIndex = 10
            ElseIf intClassPercent >= 79 AndAlso intClassPercent <= 85 Then
                'Set intClassArrayIndex variable to 11 for Warlock
                intClassArrayIndex = 11
            ElseIf intClassPercent >= 86 AndAlso intClassPercent <= 100 Then
                'Set intClassArrayIndex variable to 12 for Wizard
                intClassArrayIndex = 12
            End If

        End If

        Return intClassArrayIndex
    End Function

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Message Box to inform the purpose of this application
        MessageBox.Show("This application generates the basics of a Dungeons and Dragons 5th Edition character.",
                        "Purpose", MessageBoxButtons.OK, MessageBoxIcon.Information)

        'Fill Race and Class Selection Lists with appropriate choices and set SelectedIndex to 0 by default
        For Each strRaceType As String In strRace
            lstRaceChoice.Items.Add(strRaceType)
        Next strRaceType
        lstRaceChoice.SelectedIndex = 0

        For Each strClassType As String In strClass
            lstClassChoice.Items.Add(strClassType)
        Next strClassType
        lstClassChoice.SelectedIndex = 0

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Close the application
        Me.Close()
    End Sub

    'Functions for assigning values to each of the 6 attributes: Strength, Dexterity, Constitution, Charisma, Intelligence, Wisdom
    Private Function AssignStrength(ByVal intRaceNum As Integer)
        'Calculate Strength attribute value, taking into account race type for added benefit
        Dim intStrength As Integer
        intStrength = randRoll6.Next(1, 7) + randRoll6.Next(1, 7) + randRoll6.Next(1, 7)

        If intRaceNum = 5 OrElse intRaceNum = 8 Then
            'If race is Dragonborn or Half-Orc, add +2 to strength
            intStrength += 2
        ElseIf intRaceNum = 1 Then
            'If race is Human, add +1 to strength
            intStrength += 1
        End If
        Return intStrength
    End Function
    Private Function AssignDexterity(ByVal intRaceNum As Integer)
        'Calculate Dexterity attribute value, taking into account race type for added benefit
        Dim intDexterity As Integer
        intDexterity = randRoll6.Next(1, 7) + randRoll6.Next(1, 7) + randRoll6.Next(1, 7)

        If intRaceNum = 3 OrElse intRaceNum = 4 Then
            'If race is Elf or Halfling, add +2 to dexterity
            intDexterity += 2
        ElseIf intRaceNum = 1 OrElse intRaceNum = 7 Then
            'If race is Human or Half-Elf, add +1 to dexterity
            intDexterity += 1
        End If
        Return intDexterity
    End Function
    Private Function AssignConstitution(ByVal intRaceNum As Integer)
        'Calculate Constitution attribute value, taking into account race type for added benefit
        Dim intConstitution As Integer
        intConstitution = randRoll6.Next(1, 7) + randRoll6.Next(1, 7) + randRoll6.Next(1, 7)

        If intRaceNum = 2 Then
            'If race is Dward, add +2 to constitution
            intConstitution += 2
        ElseIf intRaceNum = 1 OrElse intRaceNum = 7 OrElse intRaceNum = 8 Then
            'If race is Human, Half-Elf, or Half-Orc, add +1 to constitution
            intConstitution += 1
        End If
        Return intConstitution
    End Function
    Private Function AssignIntelligence(ByVal intRaceNum As Integer)
        'Calculate Intelligence attribute value, taking into account race type for added benefit
        Dim intIntelligence As Integer
        intIntelligence = randRoll6.Next(1, 7) + randRoll6.Next(1, 7) + randRoll6.Next(1, 7)

        If intRaceNum = 6 Then
            'If race is Gnome, add +2 to intelligence
            intIntelligence += 2
        ElseIf intRaceNum = 1 OrElse intRaceNum = 9 Then
            'If race is Human or Tiefling, add +1 to intelligence
            intIntelligence += 1
        End If
        Return intIntelligence
    End Function
    Private Function AssignWisdom(ByVal intRaceNum As Integer)
        'Calculate Wisdom attribute value, taking into account race type for added benefit
        Dim intWisdom As Integer
        intWisdom = randRoll6.Next(1, 7) + randRoll6.Next(1, 7) + randRoll6.Next(1, 7)

        If intRaceNum = 1 Then
            intWisdom += 1
        End If
        Return intWisdom
    End Function
    Private Function AssignCharisma(ByVal intRaceNum As Integer)
        'Calculate Charisma attribute vaule, taking into account race type for added benefit
        Dim intCharisma As Integer
        intCharisma = randRoll6.Next(1, 7) + randRoll6.Next(1, 7) + randRoll6.Next(1, 7)

        If intRaceNum = 7 OrElse intRaceNum = 9 Then
            'If race is Half-Elf or Tiefling, add +2 to charisma
            intCharisma += 2
        ElseIf intRaceNum = 1 OrElse intRaceNum = 5 Then
            'if race is Human or Dragonborn, add +1 to charisma
            intCharisma += 1
        End If
        Return intCharisma
    End Function

    'Function to determine character's speed based on race
    Private Function DetermineSpeed(intRaceNum)
        Dim intSpeed As Integer

        If intRaceNum = 2 OrElse intRaceNum = 4 OrElse intRaceNum = 6 Then
            'If race is Dwarf, Halfling, or Gnome set speed to 25ft per round
            intSpeed = 25
        Else
            'Else, set speed to 30ft per round
            intSpeed = 30
        End If
        Return intSpeed
    End Function

    'Function to determine character's Hit Die based on class
    Private Function DetermineHitDie(intClassNum)
        Dim intHitDie As Integer

        If intClassNum = 1 Then
            'If class is Barbarian, set hitdie to d12
            intHitDie = 12
        ElseIf intClassNum = 12 OrElse intClassNum = 10 Then
            'If class is Sorcerer or Wizard, set hitdie to d6
            intHitDie = 6
        ElseIf intClassNum = 5 OrElse intClassNum = 7 OrElse intClassNum = 8 Then
            'If class is Fighter, Paladin, or Ranger, set hitdie to d10
            intHitDie = 10
        Else
            'Else set all other classes' hitdie to d8 (Bard, Cleric, Druid, Monk, Rogue, and Warlock)
            intHitDie = 8
        End If
        Return intHitDie
    End Function

    'Function to assign the character a background
    Private Sub ChooseBackground()
        Dim randRoll8 As New Random
        Dim strBackground As String() = {"Alcolyte", "Charlatan", "Criminal", "Entertainer", "Folk Hero",
            "Guild Artisan", "Hermit", "Noble", "Outlander", "Sage", "Sailor", "Soldier", "Urchin"}
        Dim intBackgroundPercent As Integer
        intBackgroundPercent = GetDiceRoll100()

        If intBackgroundPercent >= 1 AndAlso intBackgroundPercent <= 7 Then
            'Set background to Alcolyte
            'Personality Traits for Alcolytes:
            Dim strPersonality As String() = {"I idolize a particular hero of my faith, and constantly refer to that person's deeds and example.",
                                              "I can find common ground between the fiercest enemies, empathizing with them and always working toward peace.",
                                              "I see omens in every even and action. The gods try to speak to us, we just need to listen.",
                                              "Nothing can shake my optimistic attitude.",
                                              "I quote (or misquote) sacred texts and provers in almost every situation.",
                                              "I am tolerant (or intolerant) of other faiths and respect (or condemn) the worship of other gods.",
                                              "I've enjoyed find food, drink, and high society among my temple's elite. Rough living grates on me.",
                                              "I've spent so long in the temple that I have little practical experience dealing with people in the outside world."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Alcolytes:
            Dim strIdeal As String() = {"Tradition. The ancient traditions of worhip and sacrifice must be preserved and upheld.",
                                               "Charity. I always try to help those in need, no matter what the personal cost.",
                                               "Change. We must help bring about the changes the gods are constantly working in the world.",
                                               "Power. I hope to one day rise to the top of my faith's religious hierarchy.",
                                               "Faith. I trust that my deity will guide my actions. I have faith that if I work hard, things will go well.",
                                               "Aspiration. I seek to prove myself worth of my god's favor by matching my actions against his or her teachings."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Alcolytes:
            Dim strBond As String() = {"I would die to recover an ancient relic of my faith that was lost long ago.",
                                               "I will someday get revenge on the corrupt temple hierarchy who branded me a heretic.",
                                               "I owe my life to the priest who took me in when my parents died.",
                                               "Everything I do is for the common people.",
                                               "I will do anything to protext the temple where I served.",
                                               "I seek to preserve a sacred text that my enemies consider heretical and seek to destroy."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Alcolytes:
            Dim strFlaw As String() = {"I judge others harshly, and myself even more severely.",
                                               "I put too much trust in those who weild power within my temple's hierarchy.",
                                               "My piety sometimes leads me to blindly trust those that profess faith in my god.",
                                               "I am inflexible in my thinking.",
                                               "I am suspicious of strangers and expect the worst of them.",
                                               "Once I pick a goal, I become obsessed with it to the detriment of everything else in my life."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(0)
        ElseIf intBackgroundPercent >= 8 AndAlso intBackgroundPercent <= 14 Then
            'Set background to Charlatan    
            'Personality Traits for Charlatans:
            Dim strPersonality As String() = {"I fall in and out of love easily, and am always pursuing someone.",
                                                       "I have a joke for every occasion, especially occasions where humor is inappropriate.",
                                                       "Flattery is my preferred trick for getting what I want.",
                                                       "I'm a born gambler who can't resist taking a risk for a potential payoff.",
                                                       "I lie about almost everything, even when there's no good reason to.",
                                                       "Sarcasm and insults are my weapons of choice.",
                                                       "I keep multiple holy symbols on me and invoke whatever deity might come in useful at any given moment.",
                                                       "I pocket anything I see that might have some value."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Charlatans:
            Dim strIdeal As String() = {"Independence. i am a free spirit -- no one tells me what to do.",
                                                 "Fairness. I never target people who can't afford to lose a few coins.",
                                                 "Charity. I distribute the money I acquire to the people who really need it.",
                                                 "Creativity. I never run the same con twice.",
                                                 "Friendship. Material goods come and go. Bonds of friendship last forever.",
                                                 "Aspiration. I'm determined to make something of myself."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Charlatans:
            Dim strBond As String() = {"I fleeced the wrong person and must work to ensure that this individual never crosses paths with me or those I care about.",
                                                "I owe everything to my mentor -- a horrible person who's probably rotting in jail somewhere.",
                                                "Somewhere out there, I have a child who doesn't know me. I'm making the world better for him or her.",
                                                "I come form a noble family, and one day I'll reclaim my lands and title from those who stole them from me.",
                                                "A powerful person killed someone I love. Some day soon, I'll have my revenge.",
                                                "I swindled and ruined a person who didn't deserve it. I seek to atone for my misdeeds but might never be able to forgive myself."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Charlatans:
            Dim strFlaw As String() = {"I can't resist a pretty face.",
                                                "I'm always in debt. I spend my ill-gotten gains on decadent luxuries faster than I bring them in.",
                                                "I'm convinced that no one could ever fool me the way I fool others.",
                                                "I'm too greedy for my own good. I can't resist taking a risk if there's money involved.",
                                                "I can't resist swindling people who are more powerful than me.",
                                                "I hate to admit it and will hate myself for it, but I'll run and preserve my own hide if the going gets tough."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(1)
        ElseIf intBackgroundPercent >= 15 AndAlso intBackgroundPercent <= 21 Then
            'Set background to Criminal
            'Personality Traits for Criminals:
            Dim strPersonality As String() = {"I always have a plan for what to do when things go wrong.",
                                                      "I am always calm, no matter what the situation. I never raise my voice or let my emotions control me.",
                                                      "The first thing I do in a new place is note the locations of everything valuable -- or where such things could be hidden.",
                                                      "I would rather make a new friend than a new enemy.",
                                                      "I am incredibly slow to trust. Those who seem the fairest often have the most to hide.",
                                                      "I don't pay attention to the risks in a situation. Never tell me the odds.",
                                                      "The best way to get me to do something is to tell me I can't do it.",
                                                      "I blow up at the slightest insult."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Criminals:
            Dim strIdeal As String() = {"Honor. I don't steal from others in the trade.",
                                        "Freedom. Chains are meant to be broken, as are those who would forge them.",
                                        "Charity. I steal from the wealthy so that I can help people in need.",
                                        "Greed. I will do whatever it takes to become wealthy.",
                                        "People. I'm loyal to my friends, not to any ideals, and everyone else can take a trip down the Styx for all I care.",
                                        "Redemption. There's a spark of good in everyone."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Criminals:
            Dim strBond As String() = {"I'm trying to pay off an old debt I own to a generous benefactor.",
                                       "My ill-gotten gains go to support my family.",
                                       "Something important was taken from me, and I aim to steal it back.",
                                       "I will become the greatest thief that ever lived.",
                                       "I'm guilty of a terrible crime. I hope I can redeem myself for it.",
                                       "Someone I loved died because of a mistake I made. That will never happen again."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Criminals:
            Dim strFlaw As String() = {"When I see something valuable, I can't think about anything but how to steal it.",
                                       "When faced with a choice between money and my friends, I usually choose money.",
                                       "If there's a plan, I'll forget it. If I don't forget it, I'll ignore it.",
                                       "I have a 'tell' that reveals when I'm lying.",
                                       "I turn tail and run when things look bad.",
                                       "An innocent person is in prison for a crime that I committed. I'm okay with that."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(2)
        ElseIf intBackgroundPercent >= 22 AndAlso intBackgroundPercent <= 28 Then
            'Set background to Entertainer
            'Personality Traits for Entertainer:
            Dim strPersonality As String() = {"I know a story relevant to almost every situation.",
                                              "Whenever I come to a new place, I collect local rumors and spread gossip.",
                                              "I'm a hopeless romantic, always searching for that 'special someone.'",
                                              "Nobody stays angry at me or around me for long, since I can defuse any amount of tension.",
                                              "I love a good insult, even one directed at me.",
                                              "I get bitter if I'm not the center of attention.",
                                              "I'll settle for nothign less than perfection.",
                                              "I change my mood or my mind as quickly as i change key in a song."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Entertainer:
            Dim strIdeal As String() = {"Beauty. When I perform, I make the world better than it was.",
                                        "Tradition. The sotries, legends, and songs of the past must never be forgotten, for they teach us who we are.",
                                        "Creativity. The world is in need of new ideas and bold action.",
                                        "Greed. I'm only in it for the money and fame.",
                                        "People. I like seeing the smiles on people's faces when I perform. That's all that matters.",
                                        "Honesty. Art should reflect the soul; it should come from within and reveal who we really are."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Entertainer:
            Dim strBond As String() = {"My instrument is my most treasured possession, and it remindes me of someone I love.",
                                       "Someone stole my precious instrument, and someday I'll get it back.",
                                       "I want to be famous, whatever it takes.",
                                       "I idolize a hero of the old tales and measure my deeds against that person's.",
                                       "I will do anything to prove myself superior to my hated rival.",
                                       "I would do anything for the other members of my old troupe."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Entertainer:
            Dim strFlaw As String() = {"I'll do anything to win fame and renown.",
                                       "I'm a sucker for a pretty face.",
                                       "A scandal prevents me from ever going home again. That kind of trouble seems to follow me around.",
                                       "I once satirized a noble who still wants my head. It was a mistake that I will likely repeat.",
                                       "I have trouble keeping my true feelings hidden. My sharp tounge lands me in trouble.",
                                       "Despite my best efforts, I am unreliable to my friends."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(3)
        ElseIf intBackgroundPercent >= 29 AndAlso intBackgroundPercent <= 36 Then
            'Set background to Folk Hero
            'Personality Traits for Folk Hero:
            Dim strPersonality As String() = {"I judge people by their actions, not their words.",
                                              "If someone is in trouble, I'm always ready to lend help.",
                                              "When I set my mind to something, I follow through no matter what gets in my way.",
                                              "I have a strong sense of fair play and always try to find the most equitable solution to arguments.",
                                              "I'm confident in my own abilities and do what I can to instill confidence in others.",
                                              "Thinking is for other people. I prefer action.",
                                              "I misuse long words in an attempt to sound smarter.",
                                              "I get bored easily. When am I going to get on with my destiny?"}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Folk Hero:
            Dim strIdeal As String() = {"Respect. People deserve to be treated with dignity and respect.",
                                        "Fairness. No one should get preferential treatment.",
                                        "Freedom. Tyrants must not be allowed to oppress the people.",
                                        "Might. If I become strong, I can take what I want -- what I deserve.",
                                        "Sincerity. There's no good in pretending to be something I'm not.",
                                        "Destiny. Nothing and no one can steer me away from my higher calling."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Folk Hero:
            Dim strBond As String() = {"I have a family, but I have no idea where they are. One day, I hope to see them again.",
                                       "I worked the land, I love the land, and I will protect the land.",
                                       "A proud noble once gave me a horrible beating, and I will take my revenge on any bully I encounter.",
                                       "My tools are symbols of my past life, and I carry them so that I will never forget my roots.",
                                       "I protect those who cannot protect themselves.",
                                       "I wish my childhood sweetheart had come with me to pursue my destiny."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Folk Hero:
            Dim strFlaw As String() = {"The tyrant who rules my land will stop at nothing to see me killed.",
                                       "I'm convinced of the significance of my destiny, and blind to my shortcomings and the risk of failure.",
                                       "The people who knew me when I was young know my shameful secrect, so I can never go home again.",
                                       "I have a weakness for the vices of the city, especially hard drink.",
                                       "Secrectly, I believe that things would be better if I were a tryant lording over the land.",
                                       "I have trouble trusting in my allies."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(4)
        ElseIf intBackgroundPercent >= 37 AndAlso intBackgroundPercent <= 44 Then
            'Set background to Guild Artisan
            'Personality Traits for Guild Artisan:
            Dim strPersonality As String() = {"I believe that anything worth doing is worth doing right. I can't help it -- I'm a perfectionist.",
                                              "I'm a snob who looks down on those who can't appreciate fine art.",
                                              "I always want to know how things work and what makes people tick.",
                                              "I'm full of witty aphorisms and have a proverb for every occasion.",
                                              "I'm rude to people who lack my commitment to hard work and fair play.",
                                              "I like to talk at length about my profession.",
                                              "I don't part with my money easily and will haggle tirelessly to get the best deal possible.",
                                              "I'm well known for my work, and I want to make sure everyone appreciates it. I'm always taken aback when people haven't heard of me."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Guild Artisan:
            Dim strIdeal As String() = {"Community. It is the duty of all civilized people to strengthen the bonds of community and the security of civilization.",
                                        "Generostiy. My talents were given to me so that I could use them to benefit the world.",
                                        "Freedom. Everyone should be free to pursue his or her own livelihood.",
                                        "Greed. I'm only in it for the money.",
                                        "People. I'm committed to the people I care about, not to ideals.",
                                        "Aspiration. I work hard to be the best there is at my craft."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Guild Artisan:
            Dim strBond As String() = {"The workshop where I learned my trade is the most important place in the world to me.",
                                       "I created a great work for someone, and then found them unworthy to receive it. I'm still looking for someone worthy.",
                                       "I owe my guild a great debt for forging me into the person I am today.",
                                       "I pursue wealth to secure someone's love.",
                                       "One day I will return to my guild and prove that I am the greatest artisan of them all.",
                                       "I will get revenge on the evil forces that destroyed my place of business and ruined my livelihood."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Guild Artisan:
            Dim strFlaw As String() = {"I'll do anything to get my hands on something rare or priceless.",
                                       "I'm quick to assume that someone is trying to cheat me.",
                                       "No one must ever learn that I once stole money from guild coffers.",
                                       "I'm never satisfied with that I have -- I always want more.",
                                       "I would kill to acquire a noble title.",
                                       "I'm horribly jealous of anyone who can outshine my handiwork. Everywhere I go, I'm surrounded by rivals."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(5)
        ElseIf intBackgroundPercent >= 45 AndAlso intBackgroundPercent <= 52 Then
            'Set background to Hermit
            'Personality Traits for Hermit:
            Dim strPersonality As String() = {"I've been isolated for so long that I rarely speak, preferring gestures and the occasional grunt.",
                                              "I am utterly serene, even in the face of disaster.",
                                              "The leader of my community had something wise to say on every topic, and I am eager to share that wisdom.",
                                              "I feel tremendous empathy for all who suffer.",
                                              "I'm oblivious to etiquette and social expectations.",
                                              "I connect everything that happens to me to a grand, cosmic plan.",
                                              "I often get lost in my own thoughts and contemplation, becoming oblivious to my surroundings.",
                                              "I am working on a grand philosophical theory and love sharing my ideas."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Hermit:
            Dim strIdeal As String() = {"Greater Good. My gifts are meant to be shared with all, not used for my own benefit.",
                                        "Logic. Emotions must not cloud our sense of what is right and true, or our logical thinking.",
                                        "Free Thinking. Inquiry and curiousity are the pillars of progress.",
                                        "Power. Solitude and contemplation are paths toward mystical or magical power.",
                                        "Live and Let Live. Meddling in the affairs of others only causes trouble.",
                                        "Self-Knowledge. If you know yourself, there's nothing left to know."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Hermit:
            Dim strBond As String() = {"Nothing is more important than the other members of my hermitage, order, or association.",
                                       "I entered seclusion to hide from the ones who might still be hunting me. I must someday confront them.",
                                       "I'm still seeking the enlightenment I pursued in my seclusion, and it still eludes me.",
                                       "I entered seclusion because I loved someone I could not have.",
                                       "Should my discovery come to light, it could bring ruin to the world.",
                                       "My isolation gave me great insight into a great evil that only I can destroy."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Hermit:
            Dim strFlaw As String() = {"Now that I've returned to the world, I enjoy its delights a little too much.",
                                       "I harbor dark, bloodthirsty thoughts that my isolation and meditation failed to quell.",
                                       "I am dogmatic in my thoughts and philosophy.",
                                       "I let my need to win arguments overshadow friendships and harmony.",
                                       "I'd risk too much to uncover a lost bit of knowledge.",
                                       "I like keeping secrets and won't share them with anyone."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(6)
        ElseIf intBackgroundPercent >= 53 AndAlso intBackgroundPercent <= 60 Then
            'Set background to Noble  
            'Personality Traits for Noble:
            Dim strPersonality As String() = {"My eloquent flattery makes everyone I talk to feel like the most wonderful and important person in the world.",
                                              "The common folk love me for my kindness and generosity.",
                                              "No one could doubt by looking at my regal bearing that I am cut above the unwashed masses.",
                                              "I take great pains to always look my best and follow the latest fashions.",
                                              "I don't like to get my hands dirty, and I won't be caught dead in unsuitable accommodations.",
                                              "Despite my noble birth, I do not place myself above other folk. We all have the same blood.",
                                              "My favor, once lost, is lost forever.",
                                              "If you do me an injury, I will crush you, ruin your name, and salt your fields."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Noble:
            Dim strIdeal As String() = {"Respect. Respect is due to me because of my position, but all people regardless of station deserve to be treated with dignity.",
                                        "Responsibility. It is my duty to respect the authority of those above me, just as those below me must respect mine.",
                                        "Independence. I must prove that I can handle myself without the coddling of my family.",
                                        "Power. If I can attain more power, no one will tell me what to do.",
                                        "Family. Blood runs thicker than water.",
                                        "Noble Obligation. It is my duty to protect and care for the people beneath me."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Noble:
            Dim strBond As String() = {"I will face any challenge to win the approval of my family.",
                                       "My house's alliance with another noble family must be sustained at all costs.",
                                       "Nothing is more important than the other members of my family.",
                                       "I am in love with the heir of a family that my family despises.",
                                       "My loyalty to my sovereign is unwavering.",
                                       "The common folk must see me as a hero of the people."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Noble:
            Dim strFlaw As String() = {"I secretly believe that everyone is beneath me.",
                                       "I hide a truly scandalous secret that could ruin my family forever.",
                                       "I too often hear veiled insults and threats in every word addressed to me, and I'm quick to anger.",
                                       "I have an insatiable desire for carnal pleasures.",
                                       "In fact, the world does revolve around me.",
                                       "By my words and actions, I often bring shame to my family."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(7)
        ElseIf intBackgroundPercent >= 61 AndAlso intBackgroundPercent <= 68 Then
            'Set background to Outlander
            'Personality Traits for Outlander:
            Dim strPersonality As String() = {"I'm driven by a wanderlust that led me away from home.",
                                              "I watch over my friends as if they were a litter of newborn pups.",
                                              "I once ran twenty-five miles without stopping to warn my clan of an approaching orc horde. I'd do it again if I had to.",
                                              "I have a lesson for every situation, drawn from observing nature.",
                                              "I place no stock in wealthy or well-mannered folk. Money and manners won't save you from a hungry owlbear.",
                                              "I'm always picking things up, absently fiddling with them, and sometimes accidentally breaking them.",
                                              "I feel far more comfortable around animals than people.",
                                              "I was, in fact, raised by wolves."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Outlander:
            Dim strIdeal As String() = {"Change. Life is like the seasons, in constant change, and we must change with it.",
                                        "Greater Good. It is each person's responsibility to make the most happiness for the whole tribe.",
                                        "Honor. If I dishonor myself, I dishonor my whole clan.",
                                        "Might. The strongest are meant to rule.",
                                        "Nature. The natural world is more important than all the constructs of civilization.",
                                        "Glory. I must earn glory in battle, for myself and my clan."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Outlander:
            Dim strBond As String() = {"My family, clan, or tribe is the most important thing in my life, even when they are far from me.",
                                       "An injury to the unspoiled wilderness of my home is an injury to me.",
                                       "I will bring terrible wrath down on the evildoers who destroyed my homeland.",
                                       "I am the last of my tribe, and it is up to me to ensure their names enter legend.",
                                       "I suffer awful visions of a coming disaster and will do anything to prevent it.",
                                       "It is my duty to provide children to sustain my tribe."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Outlander:
            Dim strFlaw As String() = {"I am too enarmored of ale, wine, and other intoxicants.",
                                       "There's no room for caution in a life lived to the fullest.",
                                       "I remember every insult I've received and nurse a silent resentment toward anyone who's ever wronged me.",
                                       "I am slow to trust members of other races, tribes, and societies.",
                                       "Violence is my answer to almost any challenge.",
                                       "Don't expect me to save those who can't save themselves. It is nature's way that the strong thrive and the weak perish."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(8)
        ElseIf intBackgroundPercent >= 69 AndAlso intBackgroundPercent <= 76 Then
            'Set background to Sage
            'Personality Traits for Sage:
            Dim strPersonality As String() = {"I use polysyllabic words that convey the impression of great erudition.",
                                              "I've read every book in the world's greatest libraries -- or I like to boast that I have.",
                                              "I'm used to helping out those who aren't as smart as I am, and I patiently explain anything and everything to others.",
                                              "There's nothing I like more than a good mystery.",
                                              "I'm willing to listen to every side of an argument before I make my own judgment.",
                                              "I...speak...slowly...when talking...to idiots,...which...almost...everyone...is...compared...to me.",
                                              "I am horribly, horribly awkward in social situations.",
                                              "I'm convinced that people are always trying to steal my secrets."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Sage:
            Dim strIdeal As String() = {"Knowledge. The path to power and self-improvement is through knowledge.",
                                        "Beauty. What is beautiful points us beyond itself toward what is true.",
                                        "Logic. Emotions must not cloud our logical thinking.",
                                        "No Limits. Nothing should fetter the infinite possibility inherent in all existence.",
                                        "Power. Knowledge is the path to power and domination.",
                                        "Self-Improvement. The goal of a life of study is the betterment of oneself."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Sage:
            Dim strBond As String() = {"It is my duty to protect my students.",
                                       "I have an ancient text that holds terrible secrets that must not fall into the wrong hands.",
                                       "I work to preserve a library, university, scriptorium, or monastery.",
                                       "My life's work is a series of tomes related to a specific field of lore.",
                                       "I've been searching my whole life for the answer to a certain question.",
                                       "I sold my soul for knowledge. I hope to do great deeds and win it back."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Sage:
            Dim strFlaw As String() = {"I am easily distracted by the promise of information.",
                                       "Most people scream and run when they see a demon. I stop and take notes on its anatomy.",
                                       "Unlocking an ancient mystery is worth the price of a civilization.",
                                       "I overlook obvious solutions in favor of complicated ones.",
                                       "I speak without really thinking through my words, invariably insulting others.",
                                       "I can't keep a secret to save my life, or anyone else's."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(9)
        ElseIf intBackgroundPercent >= 77 AndAlso intBackgroundPercent <= 84 Then
            'Set background to Sailor
            'Personality Traits for Sailor:
            Dim strPersonality As String() = {"My friends know they can rely on me, no matter what.",
                                              "I work hard so that I can play hard when the work is done.",
                                              "I enjoy sailing into new ports and making new friends over a flagon of ale.",
                                              "I stretch the truth for the sake of a good story.",
                                              "To me, a tavern brawl is a nice way to get to know a new city.",
                                              "I never pass up a friendly wager.",
                                              "My language is as foul as an otyugh nest.",
                                              "I like a job well done, especially if I can convince someone else to do it."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Sailor:
            Dim strIdeal As String() = {"Respect. The thing that keeps a ship together is mutual respect between captain and crew.",
                                        "Fairness. We all do the work, so we all share in the rewards.",
                                        "Freedom. The sea is freedom -- the freedom to go anywhere and do anything.",
                                        "Mastery. I'm a predator, and the other ships on the sea are my prey.",
                                        "People. I'm committed to my crewmates, not to ideals.",
                                        "Aspiration. Someday I'll own my own ship and chart my own destiny."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Sailor:
            Dim strBond As String() = {"I'm loyal to my captain first, everything else second.",
                                       "The ship is most important -- crewmates and captains come and go.",
                                       "I'll always remember my first ship.",
                                       "In a harbor town, I have a paramour whose eyes nearly stole me from the sea.",
                                       "I was cheated out of my fair share of the profits, and I want to get my due.",
                                       "Ruthless pirates murdered my captain and crewmates, plundered our ship, and left me to die. Vengeance will be mine."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Sailor:
            Dim strFlaw As String() = {"I follow orders, even if I think they're wrong.",
                                       "I'll say anything to avoid having to do extra work.",
                                       "Once someone questions my courage, I never back down no matter how dangerous the situation.",
                                       "Once I start drinking, it's hard for me to stop.",
                                       "I can't help but pocket loose coins and other trinkets I come across.",
                                       "My pride will probably lead to my destruction."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(10)
        ElseIf intBackgroundPercent >= 85 AndAlso intBackgroundPercent <= 92 Then
            'Set background to Soldier
            'Personality Traits for Soldier:
            Dim strPersonality As String() = {"I'm always polite and respectful.",
                                              "I'm haunted by memories of war. I can't get the images of violence out of my mind.",
                                              "I've lost too many friends, and I'm slow to make new ones.",
                                              "I'm full of inspiring and cautionary tales from my military experience relevant to almost every combat situation.",
                                              "I can stare down a hell hound without flinching.",
                                              "I enjoy being strong and like breaking things.",
                                              "I have a crude sense of humor.",
                                              "I face problems head-on. A simple, direct solution is the best path to success."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Soldier:
            Dim strIdeal As String() = {"Greater Good. Our lot is to lay down our lives in defense of others.",
                                        "Responsibility. I do what I must and obey just authority.",
                                        "Independence. When people follow orders blindly, they embrace a kind of tyranny.",
                                        "Might. In life as in war, the stronger force wins.",
                                        "Live and Let Live. Ideals aren't worth killing over or going to war for.",
                                        "Nation. My city, nation, or people are all that matter."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Soldier:
            Dim strBond As String() = {"I would still lay down my life for the people I served with.",
                                       "Someone saved my life on the battlefield. To this day, I will never leave a friend behind.",
                                       "My honor is my life.",
                                       "I'll never forget the crushing defeat my company suffered or the enemies who dealt it.",
                                       "Those who fight beside me are those worth dying for.",
                                       "I fight for those who cannot fight for themselves."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Soldier:
            Dim strFlaw As String() = {"The monstrous enemy we faced in battle still leaves me quivering with fear.",
                                       "I have little respect for anyone who is not a proven warrior.",
                                       "I made a terrible mistake in battle that cost many lives -- and I would do anything to keep that mistake a secret.",
                                       "My hatred of my enemies is blind and unreasoning.",
                                       "I obey the law, even if the law causes misery.",
                                       "I'd rather eat my armor than admit when I'm wrong."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(11)
        ElseIf intBackgroundPercent >= 93 AndAlso intBackgroundPercent <= 100 Then
            'Set background to Urchin
            'Personality Traits for Urchin:
            Dim strPersonality As String() = {"I hide scraps of food and trinkets away in my pockets.",
                                              "I ask a lot of questions.",
                                              "I like to squeeze into small places where no one else can get to me.",
                                              "I sleep with my back to a wall or tree, with everything I own wrapped in a bundle in my arms.",
                                              "I eat like a pig and have bad manners.",
                                              "I think anyone who's nice to me is hiding evil intent.",
                                              "I don't like to bathe.",
                                              "I bluntly say what other people are hinting at or hiding."}
            lblPersonalityDisplay.Text = strPersonality(randRoll8.Next(1, 9) - 1)
            'Ideals for Urchin:
            Dim strIdeal As String() = {"Respect. All people, rich or poor, deserve respect.",
                                        "Community. We have to take care of each other, because no one else is going to do it.",
                                        "Change. The low are lifted up, and the high and mighty are brought down. Change is the nature of things.",
                                        "Retribution. The rich need to be shown what life and death are like in the gutters.",
                                        "People. I help the people who help me -- that's what keeps us alive.",
                                        "Aspiration. I'm going to prove that I'm worthy of a better life."}
            lblIdealDisplay.Text = strIdeal(randRoll6.Next(1, 7) - 1)
            'Bonds for Urchin:
            Dim strBond As String() = {"My town or city is my home, and I'll fight to defend it.",
                                       "I sponsor an orphanage to keep others from enduring what I was forced to endure.",
                                       "I owe my survival to another urchin who taught me to live on the streets.",
                                       "I owe a debt I can never repay to the person who took pity on me.",
                                       "I escaped my life of poverty by robbing an important person, and I'm wanted for it.",
                                       "No one else should have to endure the hardships I've been through."}
            lblBondDisplay.Text = strBond(randRoll6.Next(1, 7) - 1)
            'Flaws for Urchin:
            Dim strFlaw As String() = {"If I'm outnumbered, I will run away from a fight.",
                                       "Gold seems like a lot of money to me, and I'll do just about anything for more of it.",
                                       "I will never fully trust anyone other than myself.",
                                       "I'd rather kill someone in their sleep than fight fair.",
                                       "It's not stealing if I need it more than someone else.",
                                       "People who can't take care of themselves get what they deserve."}
            lblFlawDisplay.Text = strFlaw(randRoll6.Next(1, 7) - 1)
            lblBackgroundDisplay.Text = strBackground(12)
        End If
    End Sub

    'Function to determine how much starting gold pieces the character has
    Private Function AssignStartGold(intClassType)
        Dim randRoll4 As New Random
        Dim intGoldAmount As Integer

        'If selection structure to determine which equation to use, based on assigned class
        If intClassType = 1 OrElse intClassType = 4 Then
            'If classtype is Barbarian or Druid, then starting gold is 2d4x10gp
            intGoldAmount = (randRoll4.Next(1, 5) + randRoll4.Next(1, 5)) * 10

        ElseIf intClassType = 10 Then
            'If classtype is Sorcerer, then starting gold is 3d4x10gp
            intGoldAmount = (randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5)) * 10

        ElseIf intClassType = 9 OrElse intClassType = 11 OrElse intClassType = 12 Then
            'If classtype is Rogue, Warlock, or Wizard, then starting gold is 4d4x10gp
            intGoldAmount = (randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5)) * 10

        ElseIf intClassType = 2 OrElse intClassType = 3 OrElse intClassType = 5 OrElse intClassType = 7 OrElse intClassType = 8 Then
            'If classtype is Bard, Cleric, Fighter, Paladin, or Ranger, then starting gold is 5d4x10gp
            intGoldAmount = (randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5)) * 10

        ElseIf intClassType = 6 Then
            'If classtype is Monk, then starting gold is 5d4 gp
            intGoldAmount = (randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5) + randRoll4.Next(1, 5))

        End If

        Return intGoldAmount
    End Function

    'Function to determine proficiency based off level
    Private Function DetermineProficiency(intLevelNum)
        Dim intProfNum As Integer

        If intLevelNum >= 1 AndAlso intLevelNum <= 4 Then
            intProfNum = 2
        ElseIf intLevelNum >= 5 AndAlso intLevelNum <= 8 Then
            intProfNum = 3
        ElseIf intLevelNum >= 9 AndAlso intLevelNum <= 12 Then
            intProfNum = 4
        End If

        Return intProfNum
    End Function

    'Function to give the character a basic name based off race, if the player doesn't pick one
    Private Function ChooseName(intRaceNum)
        Dim strFullName As String = Nothing
        Dim strFirstName As String = Nothing
        Dim strLastName As String = Nothing
        Dim randRoll As New Random

        If intRaceNum = 1 Then
            'If race is Human, assign basic human first and last names.
            'Humans have such an excess of name choices due to there being 9 different clans/ethnicities
            'Basically- a more diverse culture overall, thus more variety in names in the code as I included all the example names for each clan
            Dim strFirstChoices() = {"Aseir", "Bardeid", "Haseid", "Khemed", "Mehmen", "Sudeiman", "Zasheir", "Atala", "Ceidil", "Hama", "Jasmal", "Meilil", "Seipora", "Yasheira",
                                     "Zasheida", "Darvin", "Dorn", "Evendur", "Gorstag", "Grim", "Helm", "Malark", "Morn", "Randal", "Stedd", "Arveene", "Esvele", "Jhessail", "Kerri",
                                     "Lureene", "Miri", "Rowan", "Shandri", "Tessele", "Bor", "Fodel", "Glar", "Grigor", "Igan", "Ivor", "Kosef", "Mival", "Orel", "Pavel",
                                     "Sergor", "Alethra", "Kara", "Katernin", "Mara", "Natali", "Olma", "Tana", "Zora", "Ander", "Blath", "Bran", "Frath", "Geth", "Lander",
                                     "Luth", "Malcer", "Stor", "Taman", "Urth", "Amafrey", "Betha", "Cefrey", "Kethra", "Mara", "Olga", "Silifrey", "Westra", "Aoth", "Bareris", "Ehput-Ki",
                                     "Kethoth", "Mumed", "Ramas", "So-Kehur", "Thazar-De", "Urhur", "Arizima", "Chathi", "Nephis", "Nulara", "Murithi", "Sefris", "Thola", "Umara", "Zolis",
                                     "Borivik", "Faurgar", "Jandar", "Kanithar", "Madislak", "Ralmevik", "Shaumar", "Vladislak", "Fyevarra", "Hulmarra", "Immith", "Imzel", "Navarra", "Shevarra",
                                     "Tammith", "Yuldra", "An", "Chen", "Chi", "Fai", "Jiang", "Jun", "Lian", "Long", "Meng", "On", "Shan", "Shui", "Tai", "Bai", "Chao", "Jia", "Lei", "Mei",
                                     "Qiao", "Shui", "Tai", "Anton", "Diero", "Marcon", "Pieron", "Rimardo", "Romero", "Salazar", "Umbero", "Balama", "Dona", "Faila", "Jalana", "Luisa", "Marta",
                                     "Quara", "Selise", "Vonda"}
            Dim strLastChoices() = {"Basha", "Dumein", "Jassan", "Khalid", "Mostana", "Pashar", "Rein", "Amblecrown", "Buckman", "Dundragon", "Evenwood", "Greycastle",
                                    "Tallstag", "Bersk", "Chernin", "Dotsk", "Kulenov", "Marsk", "Nemetsk", "Shemov", "Starag", "Brightwood", "Helder", "Hornraven",
                                    "Lackman", "Stormwind", "Windrivver", "Ankhalab", "Anskuld", "Fezim", "Hahpet", "Nathandem", "Sepret", "Uuthrakt", "Chergoba", "Dyernina",
                                    "Iltazyara", "Murnyethara", "Stayanoga", "Ulmokina", "Chien", "Huang", "Kao", "Kung", "Lao", "Ling", "Mei", "Pin", "Shin", "Sum", "Tan", "Wan",
                                    "Agosto", "Astorio", "Calabra", "Domine", "Falone", "Marivaldi", "Pisacar", "Ramondo"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 2 Then
            'If race is Dwarf, assign basic dwarven first and last names
            Dim strFirstChoices() = {"Adrik", "Alberich", "Baern", "Barendd", "Brottor", "Bruenor", "Dain", "Darrak", "Delg", "Eberk", "Einkil", "Fargrim", "Flint",
                                     "Gardain", "Harbek", "Kildrak", "Morgran", "Orsik", "Oskar", "Rangrim", "Rurik", "Taklinn", "Thorandin", "Thorin", "Tordek",
                                     "Traubon", "Travok", "Ulfgar", "Veit", "Vondal", "Amber", "Artin", "Audhild", "Bardryn", "Dagnal", "Diesa", "Eldeth", "Falkrunn",
                                     "Finellen", "Gunnloda", "Gurdis", "Helja", "Hlin", "Kathra", "Kristryd", "Ilde", "Liftrasa", "Mardred", "Riswynn", "Sannl", "Torbera", "Torgga", "Vistra"}
            Dim strLastChoices() = {"Balderk", "Battlehammer", "Brawnanvil", "Dankil", "Fireforge", "Frostbeard", "Gorunn", "Holderhek", "Ironfist", "Loderr", "Lutgehr", "Rumnaheim",
                                    "Strakeln", "Torunn", "Ungart"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 3 Then
            'If race is Elf, assign basic elven first and last names
            Dim strFirstChoices() = {"Ara", "Bryn", "Del", "Eryn", "Faen", "Innil", "Lael", "Mella", "Naill", "Naeris", "Phann", "Rael", "Rinn", "Sai", "Syllin", "Thia", "Vall",
                                     "Adran", "Aelar", "Aramil", "Arannis", "Aust", "Beiro", "Berrian", "Carric", "Enialis", "Erdan", "Erevan", "Galinndan", "Hadarai",
                                     "Heian", "Himo", "Immeral", "Ivellios", "Laucian", "Mindartis", "Paelias", "Peren", "Quarion", "Riardon", "Rolen", "Soveliss", "Thamior",
                                     "Tharivol", "Theren", "Varis", "Adrie", "Althaea", "Anastrianna", "Andraste", "Antinua", "Bethrynna", "Birel", "Caelynn", "Drusilia", "Enna",
                                     "Felosial", "Ielenia", "Jelenneth", "Keyleth", "Leshanna", "Lia", "Meriele", "Mialee", "Naivara", "Quelenna", "Quillathe", "Sariel", "Shanairra",
                                     "Shava", "Silaqui", "Theirastra", "Thia", "Vadania", "Valanthe", "Xanaphia"}
            Dim strLastChoices() = {"Amakiir", "Gemflower", "Amastacia", "Starflower", "Galanodel", "Moonwhisper", "Holimion", "Diamonddew", "Ilphelkiir", "Gemblossom",
                                    "Liadon", "Silverfrond", "Meliamne", "Oakenheel", "Nailo", "Nightbreeze", "Siannodel", "Moonbrook", "Xiloscient", "Goldpetal"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 4 Then
            'If race is Halfling, assign basic halfling first and last names
            Dim strFirstChoices() = {"Alton", "Ander", "Cade", "Corrin", "Eldon", "Errich", "Finnan", "Garret", "Lindal", "Lyle", "Merric", "Milo", "Osborn", "Perrin",
                                     "Reed", "Roscoe", "Wellby", "Andry", "Bree", "Callie", "Cora", "Euphemia", "Jillian", "Kithri", "Lavinia", "Lidda", "Merla", "Nedda",
                                     "Paela", "Portia", "Seraphina", "Shaena", "Trym", "Vani", "Verna"}
            Dim strLastChoices() = {"Brushgather", "Goodbarrel", "Greenbottle", "High-hill", "Hilltopple", "Leagallow", "Tealeaf", "Thorngage", "Tosscobble", "Underbough"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 5 Then
            'If race is Dragonborn, assign basic dragonborn first and last names
            Dim strFirstChoices() = {"Arjhan", "Balasar", "Bharash", "Donnar", "Ghesh", "Heskan", "Kriv", "Medrash", "Mehen", "Nadarr", "Pandjed", "Patrin", "Rhogar",
                                     "Shamash", "Shedinn", "Tarhun", "Torinn", "Akra", "Biri", "Daar", "Farideh", "Harann", "Havilar", "Jheri", "Kava", "Korinn", "Mishann",
                                     "Nala", "Perra", "Raiann", "Sora", "Surina", "Thava", "Uadjit"}
            Dim strLastChoices() = {"Clethtinthiallor", "Daardendrian", "Delmirev", "Drachedandion", "Fenkenkabradon", "Kepeshkmolik", "Kerrhylon", "Kimbatuul", "Linxakasendalor",
                                    "Myastan", "Nemmonis", "Norixius", "Ophinshtalajiir", "Prexijandilin", "Shestendeliath", "Turnuroth", "Verthisathurgiesh", "Yarjerit"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 6 Then
            'If race is Gnome, assign basic gnome first and last names
            Dim strFirstChoices() = {"Alston", "Alvyn", "Boddynock", "Brocc", "Burgell", "Dimble", "Eldon", "Erky", "Fonkin", "Frug", "Gerbo", "Gimble", "Glim", "Jebeddo",
                                     "Kellen", "Namfoodle", "Orryn", "Roondar", "Seebo", "Sindri", "Warryn", "Wrenn", "Zook", "Bimpnottin", "Breena", "Caramip", "Carlin",
                                     "Donella", "Duvamil", "Ella", "Ellyjobell", "Ellywick", "Lilli", "Loopmottin", "Lorilla", "Mardnab", "Nissa", "Nyx", "Oda", "Orla",
                                     "Roywyn", "Shamil", "Tana", "Waywocket", "Zanna"}
            Dim strLastChoices() = {"Beren", "Daergel", "Folkor", "Garrick", "Nackle", "Murnig", "Ningel", "Raulnor", "Scheppen", "Timbers", "Turen"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 7 Then
            'If race is Half-Elf, assign basic half-elf first and last names
            'Half-Elves are half human/half elf, and take either names, so this will also be long
            Dim strFirstChoices() = {"Aseir", "Bardeid", "Haseid", "Khemed", "Mehmen", "Sudeiman", "Zasheir", "Atala", "Ceidil", "Hama", "Jasmal", "Meilil", "Seipora", "Yasheira",
                                     "Zasheida", "Darvin", "Dorn", "Evendur", "Gorstag", "Grim", "Helm", "Malark", "Morn", "Randal", "Stedd", "Arveene", "Esvele", "Jhessail", "Kerri",
                                     "Lureene", "Miri", "Rowan", "Shandri", "Tessele", "Bor", "Fodel", "Glar", "Grigor", "Igan", "Ivor", "Kosef", "Mival", "Orel", "Pavel",
                                     "Sergor", "Alethra", "Kara", "Katernin", "Mara", "Natali", "Olma", "Tana", "Zora", "Ander", "Blath", "Bran", "Frath", "Geth", "Lander",
                                     "Luth", "Malcer", "Stor", "Taman", "Urth", "Amafrey", "Betha", "Cefrey", "Kethra", "Mara", "Olga", "Silifrey", "Westra", "Aoth", "Bareris", "Ehput-Ki",
                                     "Kethoth", "Mumed", "Ramas", "So-Kehur", "Thazar-De", "Urhur", "Arizima", "Chathi", "Nephis", "Nulara", "Murithi", "Sefris", "Thola", "Umara", "Zolis",
                                     "Borivik", "Faurgar", "Jandar", "Kanithar", "Madislak", "Ralmevik", "Shaumar", "Vladislak", "Fyevarra", "Hulmarra", "Immith", "Imzel", "Navarra", "Shevarra",
                                     "Tammith", "Yuldra", "An", "Chen", "Chi", "Fai", "Jiang", "Jun", "Lian", "Long", "Meng", "On", "Shan", "Shui", "Tai", "Bai", "Chao", "Jia", "Lei", "Mei",
                                     "Qiao", "Shui", "Tai", "Anton", "Diero", "Marcon", "Pieron", "Rimardo", "Romero", "Salazar", "Umbero", "Balama", "Dona", "Faila", "Jalana", "Luisa", "Marta",
                                     "Quara", "Selise", "Vonda", "Ara", "Bryn", "Del", "Eryn", "Faen", "Innil", "Lael", "Mella", "Naill", "Naeris", "Phann", "Rael", "Rinn", "Sai", "Syllin", "Thia", "Vall",
                                     "Adran", "Aelar", "Aramil", "Arannis", "Aust", "Beiro", "Berrian", "Carric", "Enialis", "Erdan", "Erevan", "Galinndan", "Hadarai",
                                     "Heian", "Himo", "Immeral", "Ivellios", "Laucian", "Mindartis", "Paelias", "Peren", "Quarion", "Riardon", "Rolen", "Soveliss", "Thamior",
                                     "Tharivol", "Theren", "Varis", "Adrie", "Althaea", "Anastrianna", "Andraste", "Antinua", "Bethrynna", "Birel", "Caelynn", "Drusilia", "Enna",
                                     "Felosial", "Ielenia", "Jelenneth", "Keyleth", "Leshanna", "Lia", "Meriele", "Mialee", "Naivara", "Quelenna", "Quillathe", "Sariel", "Shanairra",
                                     "Shava", "Silaqui", "Theirastra", "Thia", "Vadania", "Valanthe", "Xanaphia"}
            Dim strLastChoices() = {"Basha", "Dumein", "Jassan", "Khalid", "Mostana", "Pashar", "Rein", "Amblecrown", "Buckman", "Dundragon", "Evenwood", "Greycastle",
                                    "Tallstag", "Bersk", "Chernin", "Dotsk", "Kulenov", "Marsk", "Nemetsk", "Shemov", "Starag", "Brightwood", "Helder", "Hornraven",
                                    "Lackman", "Stormwind", "Windrivver", "Ankhalab", "Anskuld", "Fezim", "Hahpet", "Nathandem", "Sepret", "Uuthrakt", "Chergoba", "Dyernina",
                                    "Iltazyara", "Murnyethara", "Stayanoga", "Ulmokina", "Chien", "Huang", "Kao", "Kung", "Lao", "Ling", "Mei", "Pin", "Shin", "Sum", "Tan", "Wan",
                                    "Agosto", "Astorio", "Calabra", "Domine", "Falone", "Marivaldi", "Pisacar", "Ramondo", "Amakiir", "Gemflower", "Amastacia", "Starflower", "Galanodel", "Moonwhisper", "Holimion", "Diamonddew", "Ilphelkiir", "Gemblossom",
                                    "Liadon", "Silverfrond", "Meliamne", "Oakenheel", "Nailo", "Nightbreeze", "Siannodel", "Moonbrook", "Xiloscient", "Goldpetal"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = strLastChoices(randRoll.Next(1, strLastChoices.Length))
        ElseIf intRaceNum = 8 Then
            'If race is Half-Orc, assign basic half-orc first and last names
            'Half-Orcs don't have last names
            Dim strFirstChoices() = {"Dench", "Feng", "Gell", "Henk", "Holg", "Imsh", "Keth", "Krusk", "Mhurren", "Ront", "Shump", "Thokk", "Baggi", "Emen", "Engong", "Kansif", "Myev",
                                     "Neega", "Ovak", "Ownka", "Shautha", "Sutha", "Vola", "Volen", "Yevelda"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = Nothing
        ElseIf intRaceNum = 9 Then
            'If race is Tiefling, assign basic tiefling first and last names
            'Teiflings don't have last names
            Dim strFirstChoices() = {"Akmenos", "Amnon", "Barakas", "Damakos", "Ekemon", "Iados", "Kairon", "Leucis", "Melech", "Mordai", "Morthos", "Pelaios", "Skamos", "Therai",
                                     "Akta", "Anakis", "Bryseis", "Criella", "Damaia", "Ea", "Kallista", "Lerissa", "Makaria", "Nemeia", "Orianna", "Phelaia", "Rieta", "Art", "Carrion",
                                     "Chant", "Creed", "Despair", "Excellence", "Fear", "Glory", "Hope", "Ideal", "Music", "Nowhere", "Open", "Poetry", "Quest", "Random", "Reverence",
                                     "Sorrow", "Temerity", "Torment", "Weary"}

            strFirstName = strFirstChoices(randRoll.Next(1, strFirstChoices.Length))
            strLastName = Nothing
        End If

        strFullName = strFirstName & " " & strLastName
        Return strFullName
    End Function

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        'Generate a character with the selected choices 'bulk of the code right here.
        'Declare Variables
        'Variables for input purposes:
        Dim strPlayerName As String
        Dim strCharacterName As String
        'Variables for storing generated values:
        Dim intRaceChoice As Integer
        Dim intClassChoice As Integer
        Dim intStrength As Integer
        Dim intDexterity As Integer
        Dim intConstitution As Integer
        Dim intIntelligence As Integer
        Dim intWisdom As Integer
        Dim intCharisma As Integer
        'This one is mostly here to clarify the generated character is set to base level 1.
        'Most characters start at level 1, however it is possible to create a higher level character, as long as it is low level (2,3)
        'This is here and shown on frmMain for clarification, and possibly to later implement choice of starting level?
        Dim intLevel As Integer = 1


        '--------------------------------------------------------------------------------------
        'Determine if Player Name was input, if not, don't allow generation
        If txtPlayerName.Text.Trim = Nothing Then
            'If user did not input Player Name into txtPlayerName, inform them that they must to continue
            MessageBox.Show("Please enter the player's name.", "Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            'If the user entered their name in txtPlayerName, set strPlayerName to that
            strPlayerName = txtPlayerName.Text

            'Work with selected race and class choices
            'If no specific race was picked, generate one
            If lstRaceChoice.SelectedIndex = 0 Then
                intRaceChoice = ChooseRace()
            Else
                intRaceChoice = lstRaceChoice.SelectedIndex
            End If
            'If no specific class was picked, generate one
            If lstClassChoice.SelectedIndex = 0 Then
                intClassChoice = ChooseClass(intRaceChoice)
            Else
                intClassChoice = lstClassChoice.SelectedIndex
            End If

            'If txtCharName is empty, choose a basic name for the character based on class. 
            If txtCharName.Text.Trim = Nothing Then
                strCharacterName = ChooseName(intRaceChoice)
            Else
                strCharacterName = txtCharName.Text
            End If

            'Assign values to attribute variables
            intStrength = AssignStrength(intRaceChoice)
            intDexterity = AssignDexterity(intRaceChoice)
            intConstitution = AssignConstitution(intRaceChoice)
            intIntelligence = AssignIntelligence(intRaceChoice)
            intWisdom = AssignWisdom(intRaceChoice)
            intCharisma = AssignCharisma(intRaceChoice)

            ChooseBackground()

            '--------------------------------------------------------------------------------------
            'Output display into labels
            '   Character Info
            lblPlayerNameDisplay.Text = strPlayerName
            lblCharNameDisplay.Text = strCharacterName
            lblRaceDisplay.Text = strRace(intRaceChoice)
            lblClassDisplay.Text = strClass(intClassChoice)
            lblLevelDisplay.Text = intLevel

            '   Stats
            lblStrengthDisplay.Text = intStrength
            lblDexterityDisplay.Text = intDexterity
            lblConstitutionDisplay.Text = intConstitution
            lblIntelligenceDisplay.Text = intIntelligence
            lblWisdomDisplay.Text = intWisdom
            lblCharismaDisplay.Text = intCharisma

            lblSpeedDisplay.Text = DetermineSpeed(intRaceChoice) & "ft"
            lblHitDiceDisplay.Text = "d" & DetermineHitDie(intClassChoice)
            lblStartGoldDisplay.Text = AssignStartGold(intClassChoice) & "gp"
            lblProficiencyDisplay.Text = "+" & DetermineProficiency(intLevel)
        End If

    End Sub

    Private Sub txtCharName_GotFocus(sender As Object, e As EventArgs) Handles txtCharName.GotFocus
        txtCharName.SelectAll()
    End Sub

    Private Sub txtPlayerName_GotFocus(sender As Object, e As EventArgs) Handles txtPlayerName.GotFocus
        txtPlayerName.SelectAll()
    End Sub

    'Sub procedure to clear labels
    Private Sub ClearLabels()
        lblPlayerNameDisplay.Text = String.Empty
        lblCharNameDisplay.Text = String.Empty
        lblRaceDisplay.Text = String.Empty
        lblClassDisplay.Text = String.Empty
        lblBackgroundDisplay.Text = String.Empty
        lblLevelDisplay.Text = String.Empty
        lblStrengthDisplay.Text = String.Empty
        lblDexterityDisplay.Text = String.Empty
        lblConstitutionDisplay.Text = String.Empty
        lblIntelligenceDisplay.Text = String.Empty
        lblWisdomDisplay.Text = String.Empty
        lblCharismaDisplay.Text = String.Empty
        lblSpeedDisplay.Text = String.Empty
        lblHitDiceDisplay.Text = String.Empty
        lblProficiencyDisplay.Text = String.Empty
        lblStartGoldDisplay.Text = String.Empty
        lblPersonalityDisplay.Text = String.Empty
        lblIdealDisplay.Text = String.Empty
        lblBondDisplay.Text = String.Empty
        lblFlawDisplay.Text = String.Empty
    End Sub

    'When input objects have changed events, clear output labels
    Private Sub txtPlayerName_TextChanged(sender As Object, e As EventArgs) Handles txtPlayerName.TextChanged
        'When text is changed in txtPlayerName, clear all the labels with output from btnGenerate_Click
        ClearLabels()
    End Sub

    Private Sub txtCharName_TextChanged(sender As Object, e As EventArgs) Handles txtCharName.TextChanged
        'When text is changed in txtCharName, clear all the labels with output from btnGenerate_Click
        ClearLabels()
    End Sub

    Private Sub lstRaceChoice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstRaceChoice.SelectedIndexChanged
        'When selected index is changed in lstRaceChoice, clear all the labels with output from btnGenerate_Click
        ClearLabels()
    End Sub

    Private Sub lstClassChoice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstClassChoice.SelectedIndexChanged
        'When selected index is changed in lstClassChoice, clear all the labels with output from btnGenerate_Click
        ClearLabels()
    End Sub

    'Only allow letters (upper and lower), space, backspace, and apostrophe to be input into txtPlayerName and txtCharName
    Private Sub txtPlayerName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPlayerName.KeyPress
        If e.KeyChar Like "[A-Z]" OrElse e.KeyChar Like "[a-z]" OrElse e.KeyChar = ControlChars.Back OrElse e.KeyChar = " " OrElse e.KeyChar = "'" Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub txtCharName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCharName.KeyPress
        If e.KeyChar Like "[A-Z]" OrElse e.KeyChar Like "[a-z]" OrElse e.KeyChar = ControlChars.Back OrElse e.KeyChar = " " OrElse e.KeyChar = "'" Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
End Class
