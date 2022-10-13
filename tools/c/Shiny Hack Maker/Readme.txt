**********************************************************
*Shiny Hack Maker v1.0.1.0 {The Special C Version} ReadMe*
**********************************************************

This program is built for implementing the famous "Shiny Hack" made by Mastermind_X. Version 1.0.9 is rebuilt from scratch as I felt like the coding sucked bad - afterall this was like my first tool but now its better =). Remember for anyone wanting any special features you must request it personnally eg. translations & support for other versions and you'll have to wait XD. That also goes for bug reports.

Features
========
- Shiny Hack v2
- Shiny Hack Removal
- Shiny Flag Routine Inserter
- Patch Detection
- Logging
- Translations: French, Italian

Scripts Inserted by the compiler
================================
WildBattle:
#org 0xOffset
callasm 0xASMPointer
wildbattle 0xPokemon 0xLevel 0xItem
end

GiveEgg:
#org 0xOffset
callasm 0xASMPointer
giveegg 0xPokemon
end

GivePokemon:
#org 0xOffset
callasm 0xASMPointer
giveegg 0xPokemon 0xLevel 0xItem 0x0 0x0 0x0
end

Trainerbattle:
#org 0xOffset
callasm 0xASMPointer
trainerbattle 0x0 0xBattleNumber 0x0 0xIntroText 0xDefeatText
msgbox 0xAfterBattleText
boxset 0x2
end

- To remove Shiny Hack from a ROM, you must double click "Patched".

Greetz
======
Ash2000
Darthatron
D-Trogh
HackMew
Interdpth
Lugiale
Mastermind_X
Thethethethe

Well... And theres the readme, hope you enjoy this program and beware of bugs or odd programming stuff and feel free to let me know as soon as possible - till then good-day to you.

Enjoy!


