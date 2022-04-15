# Visual Basic DND Generator Spring 2018

## Project Description
This is a Dungeons &amp; Dragons character generator created as the final project to my Spring 2018 CSC-139 Visual Basic course at Halifax Community College.  
Visual Basic was used as was required for the course.

This project was created on April 13th, 2018 and the current version was last edited on April 21st, 2018.
The main visual basic code for this project can be found [here](../main/DDCharacterGen%20Project/Form1.vb).

This was my first real project in coding throughout my college experience that was not following tutorials step-by-step. This was also my first step into D&D character creation and all the pieces that go into that process. That said, this program by no means perfect and I would actively like to come back to this project in the future and clean it up.

Specific items of note I'd like to change in the future include:  
- Removing unnecessary fields.  
- Ensuring that the generator is inline with the official character creation process.  
- Having the option for the output to go into a character sheet ready to be printed.

The generator requires the user input the Player name before it will generate a character. The user also has the option to input a Character name and select the Race or Class if they have a preference. If these are left with the default selection, the generator will pick for the user.  
There are processes within the code that mimic the rolling of different sided dice, just like someone would in the manual creation of a D&D character.  
When generating a race or class, the selection of specific ones weigh more than others to reflect more common or rare races and class combinations. 
For example:
- Human, Dwarf, Elf, and Halfling each have a 15% chance of being selected. While DragonBorn, Gnome, Half-elf, Half-orc, and Tiefling each have an 8% chance of being selected.
- Within each race, the more common class(es) is(are) weighted heavier, with other classes being given an equal chance to each other. 
Specific numbers chosen can be observed in the flow charts below, included in the *Design Process*.

## Design Process
The following images include the initial flow chart for the project. These were used for the base of the project, which was then changed as necessary throughout the coding and testing process.

### Flow Chart Part 1
<img src="https://user-images.githubusercontent.com/101907789/163502809-d07ef10b-0440-44d5-adc6-5d142f924cc7.JPG" width="450"></img>

### Flow Chart Part 2
<img src="https://user-images.githubusercontent.com/101907789/163502823-83c85185-40e7-4c63-973a-117ebca6e6b8.JPG" width="450"></img>

The following images show the proposed GUI design for the application that were ultimately not chosen for this project - either because they were too complicated for the time limit of this project, or did not follow best practices.

### Proposed Main Form Design
<img src="https://user-images.githubusercontent.com/101907789/163504612-b02f8b26-7563-44d3-9b3b-c13ee1e27c2a.png"></img>

### Proposed Generated Form Design 1
<img src="https://user-images.githubusercontent.com/101907789/163504621-dbad104f-b579-4417-8f93-3a3af2fdb329.png" width="450"></img>

### Proposed Generated Form Design 2
<img src="https://user-images.githubusercontent.com/101907789/163504627-9d7b3cad-6053-4502-980b-a4985752a9df.png" width="450"></img>

## How to install
If you would like to run the D&D Character Generator, the .exe file can be found [here](../main/DDCharacterGen%20Project/bin/Debug/DDCharacterGen%20Project.exe). 
Otherwise, this repository should include all Visual Basic files and resources necessary to open, edit, and run the project in Visual Studio.

## Acknowledgements
I would like to thank Mrs. Debra Dickens at Halifax Community College for teaching the Spring 2018 CSC-139 Visual Basic course and providing me the knowledge and challenge neccessary for this project. 

Materials referenced include the [*D&D Player's Handbook, 5th Edition*](https://dnd.wizards.com/products/rpg_playershandbook) and the official [5th Edition Character sheets](https://dnd.wizards.com/resources/character-sheets).
