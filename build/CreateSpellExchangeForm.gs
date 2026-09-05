/**
 * Wizard Spell Exchange -- Google Form builder (Apps Script)
 * =========================================================
 *
 * WHAT THIS IS
 *   Creates the spellbook-registration Form in YOUR Google account, with every
 *   question in the exact order the SpellExchange workbook expects. Nothing to
 *   type by hand.
 *
 * WHY SCRIPTED
 *   The workbook reads the response tab BY COLUMN POSITION. One question added,
 *   removed or reordered shifts every column and quietly breaks every formula.
 *   This script is the guarantee that the order is right.
 *
 * HOW TO RUN (about 2 minutes)
 *   1. Go to https://script.google.com  ->  "New project".
 *   2. Delete the placeholder code, paste this whole file in.
 *   3. Press Run and select the function `createSpellExchangeForm`.
 *   4. First run asks you to authorize -- approve it. Run again if it stops at
 *      the consent screen.
 *   5. Open the Execution log (Ctrl+Enter). It prints the LIVE and EDIT links.
 *
 * LINKING TO THE WORKBOOK
 *   Upload build/SpellExchange.xlsx to Drive, open it as a Google Sheet, then
 *   put its id in SPREADSHEET_ID below and re-run. The id is the long string in
 *   the sheet URL between /d/ and /edit. Leave it empty to link by hand later
 *   via Responses > Link to Sheets.
 *
 *   Google ALWAYS makes its own new tab ("Form Responses 1") -- it will never
 *   write into the workbook's own "Form Responses" tab. Repoint it afterwards,
 *   and mind the order: in Sheets, DELETING a referenced sheet turns every
 *   formula into #REF! for good, while RENAMING one rewrites the formulas to
 *   follow it. So:
 *     1. rename the workbook's "Form Responses" tab to "OldResponses"
 *     2. rename Google's new tab to exactly "Form Responses"
 *     3. Ctrl+H, find `OldResponses!` replace `'Form Responses'!`,
 *        Search: all sheets, tick "Also search within formulas"
 *     4. delete "OldResponses"
 *   SETUP.md walks through this.
 *
 * ONE SETTING THIS SCRIPT CANNOT SET
 *   Apps Script has no API for response receipts. After running, open the form
 *   and set Settings > Responses > "Send responders a copy of their response"
 *   to ALWAYS. That receipt email carries the edit link players need to fix
 *   their entry after they get home -- don't skip it.
 *
 * GENERATED FILE -- do not hand-edit.
 *   Regenerate with:  python scripts/make_form_script.py
 *   Spell list: 350 leveled wizard spells, AL-legal Forgotten Realms sources.
 */

// The workbook's Google Sheets id -- responses are linked here automatically.
// Set it with:  python scripts/make_form_script.py --spreadsheet-id <id>
// The id is the part of the sheet URL between /d/ and /edit (NOT the gid).
var SPREADSHEET_ID = "";

var SUBCLASSES = ["Abjurer", "Diviner", "Evoker", "Illusionist", "Conjurer", "Enchanter", "Necromancer", "Transmuter", "Bladesinger", "Order of Scribes", "Chronurgist", "Graviturgist", "Other / not listed"];

// Spell choices by level, generated from data/wizard-spells.csv.
var SPELLS = {
  1: [
    "Absorb Elements",
    "Alarm",
    "Burning Hands",
    "Catapult",
    "Cause Fear",
    "Charm Person",
    "Chromatic Orb",
    "Color Spray",
    "Comprehend Languages",
    "Detect Magic",
    "Disguise Self",
    "Earth Tremor",
    "Expeditious Retreat",
    "False Life",
    "Feather Fall",
    "Find Familiar",
    "Fog Cloud",
    "Frost Fingers",
    "Grease",
    "Ice Knife",
    "Identify",
    "Illusory Script",
    "Jump",
    "Longstrider",
    "Mage Armor",
    "Magic Missile",
    "Protection from Evil and Good",
    "Ray of Sickness",
    "Shield",
    "Silent Image",
    "Sleep",
    "Snare",
    "Spellfire Flare",
    "Tasha's Caustic Brew",
    "Tasha's Hideous Laughter",
    "Tenser's Floating Disk",
    "Thunderwave",
    "Unseen Servant",
    "Wardaway",
    "Witch Bolt"
  ],
  2: [
    "Aganazzar's Scorcher",
    "Air Bubble",
    "Alter Self",
    "Arcane Lock",
    "Arcane Vigor",
    "Augury",
    "Battle Familiar",
    "Blindness/Deafness",
    "Blur",
    "Cloud of Daggers",
    "Continual Flame",
    "Crown of Madness",
    "Darkness",
    "Darkvision",
    "Death Armor",
    "Deryan's Helpful Homunculi",
    "Detect Thoughts",
    "Disruptive Tune",
    "Dragon's Breath",
    "Dueling Ground",
    "Dust Devil",
    "Earthbind",
    "Elminster's Elusion",
    "Enhance Ability",
    "Enlarge/Reduce",
    "Flaming Sphere",
    "Flock of Familiars",
    "Gentle Repose",
    "Gust of Wind",
    "Hold Person",
    "Invisibility",
    "Knock",
    "Levitate",
    "Locate Object",
    "Magic Mouth",
    "Magic Weapon",
    "Maximilian's Earthen Grasp",
    "Melf's Acid Arrow",
    "Mind Spike",
    "Mirror Image",
    "Misty Step",
    "Nathair's Mischief",
    "Nystul's Magic Aura",
    "Phantasmal Force",
    "Pyrotechnics",
    "Ray of Enfeeblement",
    "Rime's Binding Ice",
    "Rope Trick",
    "Scorching Ray",
    "See Invisibility",
    "Shadow Blade",
    "Shatter",
    "Skywrite",
    "Snilloc's Snowball Swarm",
    "Spider Climb",
    "Spray of Cards",
    "Suggestion",
    "Tasha's Mind Whip",
    "Uncertain Footing",
    "Warding Wind",
    "Warp Sense",
    "Web",
    "Wither and Bloom"
  ],
  3: [
    "Animate Dead",
    "Antagonize",
    "Ashardalon's Stride",
    "Bestow Curse",
    "Blink",
    "Cacophonic Shield",
    "Catnap",
    "Clairvoyance",
    "Conjure Constructs",
    "Counterspell",
    "Dispel Magic",
    "Enemies Abound",
    "Erupting Earth",
    "Fear",
    "Feign Death",
    "Fireball",
    "Flame Arrows",
    "Fly",
    "Galder's Tower",
    "Gaseous Form",
    "Glyph of Warding",
    "Haste",
    "Hypnotic Pattern",
    "Inflict Doubt",
    "Intellect Fortress",
    "Laeral's Silver Lance",
    "Leomund's Tiny Hut",
    "Life Transference",
    "Lightning Bolt",
    "Magic Circle",
    "Major Image",
    "Melf's Minute Meteors",
    "Nondetection",
    "Phantom Steed",
    "Protection from Energy",
    "Remove Curse",
    "Sending",
    "Sleet Storm",
    "Slow",
    "Speak with Dead",
    "Spirit Shroud",
    "Stinking Cloud",
    "Summon Fey",
    "Summon Lesser Demons",
    "Summon Shadowspawn",
    "Summon Undead",
    "Sylun\u00e9\u2019s Viper",
    "Thunder Step",
    "Tidal Wave",
    "Tiny Servant",
    "Tongues",
    "Vampiric Touch",
    "Wall of Sand",
    "Wall of Water",
    "Water Breathing"
  ],
  4: [
    "Arcane Eye",
    "Backlash",
    "Banishment",
    "Blight",
    "Charm Monster",
    "Confusion",
    "Conjure Minor Elementals",
    "Control Water",
    "Dimension Door",
    "Distorted Distance",
    "Divination",
    "Elemental Bane",
    "Evard's Black Tentacles",
    "Fabricate",
    "Festering Blast",
    "Fire Shield",
    "Galder's Speedy Courier",
    "Gate Seal",
    "Greater Invisibility",
    "Hallucinatory Terrain",
    "Ice Storm",
    "Leomund's Secret Chest",
    "Locate Creature",
    "Mordenkainen's Faithful Hound",
    "Mordenkainen's Private Sanctum",
    "Otiluke's Resilient Sphere",
    "Phantasmal Killer",
    "Polymorph",
    "Raulothim's Psychic Lance",
    "Sickening Radiance",
    "Spellfire Storm",
    "Spirit of Death",
    "Stone Shape",
    "Stoneskin",
    "Storm Sphere",
    "Summon Aberration",
    "Summon Construct",
    "Summon Elemental",
    "Summon Greater Demon",
    "Vitriolic Sphere",
    "Wall of Fire",
    "Watery Sphere"
  ],
  5: [
    "Alustriel's Mooncloak",
    "Animate Objects",
    "Bigby's Hand",
    "Circle of Power",
    "Cloudkill",
    "Cone of Cold",
    "Conjure Elemental",
    "Contact Other Plane",
    "Control Winds",
    "Create Spelljamming Helm",
    "Creation",
    "Danse Macabre",
    "Dawn",
    "Dominate Person",
    "Dream",
    "Enervation",
    "Far Step",
    "Geas",
    "Grave Ground",
    "Hold Monster",
    "Immolation",
    "Infernal Calling",
    "Jallarzi's Storm of Radiance",
    "Legend Lore",
    "Mislead",
    "Modify Memory",
    "Mordenkainen\u2019s Lucubration",
    "Negative Energy Flood",
    "Passwall",
    "Planar Binding",
    "Rary's Telepathic Bond",
    "Scrying",
    "Seeming",
    "Skill Empowerment",
    "Songal's Elemental Suffusion",
    "Spirit Lantern",
    "Steel Wind Strike",
    "Summon Draconic Spirit",
    "Summon Dragon",
    "Synaptic Static",
    "Telekinesis",
    "Teleportation Circle",
    "Transmute Rock",
    "Wall of Force",
    "Wall of Light",
    "Wall of Stone",
    "Waves of Exhaustion",
    "Yolande's Regal Presence"
  ],
  6: [
    "Arcane Gate",
    "Chain Lightning",
    "Circle of Death",
    "Contingency",
    "Create Homunculus",
    "Create Undead",
    "Disintegrate",
    "Drawmij's Instant Summons",
    "Elminster's Effulgent Spheres",
    "Eyebite",
    "Fizban's Platinum Shield",
    "Flesh to Stone",
    "Globe of Invulnerability",
    "Guards and Wards",
    "Investiture of Flame",
    "Investiture of Ice",
    "Investiture of Stone",
    "Investiture of Wind",
    "Magic Jar",
    "Mass Suggestion",
    "Mental Prison",
    "Move Earth",
    "Otiluke's Freezing Sphere",
    "Otto's Irresistible Dance",
    "Programmed Illusion",
    "Scatter",
    "Soul Cage",
    "Summon Fiend",
    "Sunbeam",
    "Tasha's Bubbling Cauldron",
    "Tasha's Otherworldly Guise",
    "Tenser's Transformation",
    "True Seeing",
    "Wall of Ice"
  ],
  7: [
    "Aura of Evasion",
    "Create Magen",
    "Crown of Stars",
    "Delayed Blast Fireball",
    "Draconic Transformation",
    "Dream of the Blue Veil",
    "Etherealness",
    "Finger of Death",
    "Forcecage",
    "Fractured Awareness",
    "Mirage Arcane",
    "Mordenkainen's Magnificent Mansion",
    "Mordenkainen's Sword",
    "Plane Shift",
    "Power Word Pain",
    "Prismatic Spray",
    "Project Image",
    "Reverse Gravity",
    "Reweave Fate",
    "Sequester",
    "Simbul's Synostodweomer",
    "Simulacrum",
    "Symbol",
    "Teleport",
    "Transfix",
    "Whirlwind"
  ],
  8: [
    "Abi-Dalzim's Horrid Wilting",
    "Antimagic Field",
    "Antipathy/Sympathy",
    "Befuddlement",
    "Clone",
    "Control Weather",
    "Demiplane",
    "Dominate Monster",
    "Entrancing Mirrors",
    "Holy Star of Mystra",
    "Illusory Dragon",
    "Incendiary Cloud",
    "Iron Body",
    "Lightning Ring",
    "Maddening Darkness",
    "Maze",
    "Mighty Fortress",
    "Mind Blank",
    "Moment of Prescience",
    "Power Word Stun",
    "Sunburst",
    "Telepathy"
  ],
  9: [
    "Astral Projection",
    "Blade of Disaster",
    "Detonate",
    "Foresight",
    "Gate",
    "Hindsight",
    "Imprisonment",
    "Invulnerability",
    "Mass Polymorph",
    "Meteor Swarm",
    "Power Word Kill",
    "Prismatic Wall",
    "Psychic Scream",
    "Shapechange",
    "Time Stop",
    "True Polymorph",
    "Vision of Elapsing Eons",
    "Wail of the Banshee",
    "Weird",
    "Wish"
  ]
};


function createSpellExchangeForm() {
  var form = FormApp.create('Wizard Spell Exchange -- AL Epic (Forgotten Realms)');

  form.setDescription(
    'Register what is in your wizard\'s spellbook so other wizards at the Epic ' +
    'can see what they could copy -- and so you can see what you could copy from them.\n\n' +
    'Tick only the spells you ALREADY have. Cantrips are not included: they are ' +
    'not kept in a spellbook.\n\n' +
    'Copying costs 50 GP per spell level, plus 1 Downtime Day per spell for ' +
    'spell levels 1-4 and 2 DT per spell for levels 5-9. You may only copy a ' +
    'spell of a level you can already prepare.\n\n' +
    'AL rule (ALPG p.3): you may copy from another character\'s spellbook ' +
    'immediately after a session in which you both played -- so register today. ' +
    'For this Epic the whole event counts as one session, so any wizard here may ' +
    'copy from any other wizard here.\n\n' +
    'One entry per wizard character. You will be emailed a link to edit your ' +
    'answers later.');

  form.setProgressBar(true)
      .setCollectEmail(true)
      .setAllowResponseEdits(true)
      .setLimitOneResponsePerUser(false)
      .setShuffleQuestions(false)
      .setConfirmationMessage(
        'Registered. Check your email for a link to edit this entry, and open ' +
        'the results sheet to see what you can copy and from whom.');

  // ---- Page 1: who you are -------------------------------------------------
  // ORDER IS LOAD-BEARING. The workbook expects response columns:
  //   A Timestamp | B Email | C Player | D Character | E Wizard level
  //   F Table | G Subclass | H Contact | I..Q Level 1..9 spells | R Other
  // Page breaks and section headers create no columns, so they are free.

  form.addTextItem()
      .setTitle('Player name')
      .setRequired(true);

  form.addTextItem()
      .setTitle('Character name')
      .setHelpText('One entry per wizard. Submit again for a second character.')
      .setRequired(true);

  var levels = [];
  for (var i = 1; i <= 20; i++) { levels.push(String(i)); }
  form.addListItem()
      .setTitle('Wizard level')
      .setHelpText('Your levels in the Wizard class. This sets which spell ' +
                   'levels you are allowed to copy.')
      .setChoiceValues(levels)
      .setRequired(true);

  form.addTextItem()
      .setTitle('Table number')
      .setHelpText('Which table you are playing at today.');

  form.addListItem()
      .setTitle('Wizard subclass')
      .setHelpText('Order of Scribes matters: it lets you copy ten level 1-4 ' +
                   'spells, or five level 5-9 spells, for 1 Downtime Day.')
      .setChoiceValues(SUBCLASSES);

  form.addTextItem()
      .setTitle('Contact after the event')
      .setHelpText('Discord handle or similar, so people can reach you to ' +
                   'arrange copying. Optional.');

  // ---- One page per spell level -------------------------------------------
  // A page each, rather than one enormous page: a Tier 1 wizard would otherwise
  // have to scroll past 300 irrelevant checkboxes to reach the end.
  for (var lv = 1; lv <= 9; lv++) {
    form.addPageBreakItem()
        .setTitle('Level ' + lv + ' spells')
        .setHelpText(lv === 1
          ? 'Tick every level 1 spell already in your spellbook.'
          : 'Tick every level ' + lv + ' spell already in your spellbook. ' +
            'Nothing at this level? Just press Next.');

    form.addCheckboxItem()
        .setTitle('Level ' + lv + ' spells in your spellbook')
        .setHelpText(SPELLS[lv].length + ' spells available at this level.')
        .setChoiceValues(SPELLS[lv]);
  }

  // ---- Final page ----------------------------------------------------------
  form.addPageBreakItem()
      .setTitle('Anything missing?')
      .setHelpText('Almost done.');

  form.addParagraphTextItem()
      .setTitle('Other spells not in the lists above')
      .setHelpText('One per line, or comma separated. Use this for anything ' +
                   'the lists missed -- spell names are matched by name, so ' +
                   'spell them as printed.');

  if (SPREADSHEET_ID) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID);
    Logger.log('Responses linked to spreadsheet ' + SPREADSHEET_ID);
    Logger.log('CHECK THE TAB NAME: it must be exactly "Form Responses".');
  } else {
    Logger.log('No SPREADSHEET_ID set -- link by hand: Responses > Link to Sheets.');
  }

  PropertiesService.getScriptProperties()
      .setProperty('SPELLEXCHANGE_FORM_ID', form.getId());

  Logger.log('LIVE (share this):    ' + form.getPublishedUrl());
  Logger.log('EDIT (open to tweak): ' + form.getEditUrl());
  Logger.log('');
  Logger.log('STILL TO DO BY HAND: Settings > Responses > "Send responders a ' +
             'copy of their response" = ALWAYS. Apps Script cannot set it, and ' +
             'that email is how players edit their entry from home.');
}


// Stop the form accepting responses (run by hand after the event if you like).
function closeForm() {
  var id = PropertiesService.getScriptProperties()
      .getProperty('SPELLEXCHANGE_FORM_ID');
  if (!id) { Logger.log('No stored form id -- run createSpellExchangeForm first.'); return; }
  FormApp.openById(id).setAcceptingResponses(false);
  Logger.log('Form closed.');
}


function reopenForm() {
  var id = PropertiesService.getScriptProperties()
      .getProperty('SPELLEXCHANGE_FORM_ID');
  if (!id) { Logger.log('No stored form id -- run createSpellExchangeForm first.'); return; }
  FormApp.openById(id).setAcceptingResponses(true);
  Logger.log('Form reopened.');
}
