All of this has been vibe-coded with the free version of Claude for Stable Diffusion ReForge with python 3.10 to work on my machine. I barely know how to do 'pip install'. I can't help troubleshoot anything if you have issues. If you're having trouble, your best bet is to paste the code into into a new chat with your own free account with Claude and get it to fix it to work on your machine (it will probably help to tell it what version of python you have and which flavor of Stable Diffusion you're using while doing so). Again, I have absolutely no idea what I'm doing.

This one is has proven to be a bit buggy, so make sure to periodically back up your wildcards if you use this, and its much safer to write long edits in a different program like notepad++. That said, for inter-wildcard cohesion, this program works wonders for making all your wildcards play nice.

The marquee feature of this program is the ability to double click on a __wildcard__ within the text of a document to switch to that wildcard. This feature alone makes adding tweaks a breeze compared to scrolling though alphabetized document lists for the wildcard you want to change, but this program does so much more.

Within the main text editor window there are a few more features, namely additive highlighting based on (), {} and [] to denote increased strength as well as prompt grouping. LoRAs or anything else encased in <> are highlighted in their own color as well.

Starting at the top, there's the obligatory new, open, save and save as buttons, but the rename function does something special. when a wildcard is renamed in this program, every instance of that wildcard for all wildcards in the wildcard folder are updated with the new name. This way you can rename your wildcards freely without breaking anything. It works the same way when renaming from the document list on the left.

Next theres the browse forwards and back buttons, which work the same way as an internet browser, going forwards and backwards between which wildcards you've viewed, so you can double click a __wildcard__, make your changes, then go back to the wildcard that called it and continue from there. the mouse thumb buttons also control this feature.

Undo and redo.

The clone lines button (or Ctrl+D) lets will copy and instantly paste a highlighted line, meant to be used to weigh probabilities for a line to be called.

The wrap wildcard button turns highlighted text into a wildcard by flanking it with the defined wildcard wrapper (usually __), then creates a new document with that wildcard's name in the document list, so you can instantly double click it to start creating that wildcard.

There's a basic find and replace function, followed by a search all popup, which searches through ALL of the wildcards in the wildcard folder and has a replace feature (which can sometimes be buggy :/)

The diagnose button preforms 2 main functions: it searches through all wildcards in the wildcard folder for stray wrapping characters which if present when the dynamic prompts extension is run will completely break the wildcard function. This feature finds and automatically removes any of those issues. The other function is to look for "dead end" wildcards, either a wildcard being called that isn't in your wildcards folder, or an isolated wildcard that neither calls nor is called by any other wildcards in your wildcards folder.

The LoRA button lets you adjust the strength of LoRAs in a document or across all wildcards by an increment, such as strengthening all <lora:il_contrast_slider_d1:#> instances by +0.1 for example.

Text wrap, hotkeys, settings and the tabs toggle are all pretty basic, but the reoganize button is helpful for keeping your wildcards sorted outside of the program. On the left the main text editor window, the document list is much more malleable than it is in most programs. similar to how layer folders are presented in programs like photoshop, you can create and name folders, and drag documents in and out of them. This is where the reogranize button comes in. If you press this button, the program will create subfolders and move your wildcards around to match what is seen in the document list. Inversely, you can press the import button at the top of the document list to have the document list be populated by your wildcards in their existing folder structure. the cut isolated button in this region removes wildcards that don't call or are not called by any other wildcards from the document list (the actual .txt file is uneffected).

Underneath the document viewer is the wildcards in doc window, which lists out all wildcards present in the currently open document. double clicking any of these in this window will open that wildcard. It also features an open all button, as well as a search use button, which opens a popup that will list out all wildcards in the wildcard folder that call the currently open document.

Lastly in the bottom right corner there's a wrapper confirmation and a spell check toggle. Overall I find this tool incredibly useful, but it definitely is the buggiest out of the lot.

UPDATE: Changed the core of the text editing section to a MUCH more reliable scintilla base.
