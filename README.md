# Ash’s Engineering Drawing Maker  
### Because paying Autodesk to place a rectangle around a picture can get absolutely fucked.

---

## What The Hell Is This?

This is a **local, offline, no-licence, no-login, no-corporate-nonsense** tool for turning images and PDFs into **actual engineering drawings** without:

- AutoCAD
- SolidWorks
- Revit
- Inventor
- A subscription
- A licence server
- A phone-home DRM tantrum
- Or some MBA deciding your file is “read-only” this month

You give it images or PDFs.  
It gives you a **clean, ISO-style drawing pack PDF**.

That’s it.  
No hoops. No begging. No kneeling.

---

## Why This Exists (AKA: Corporate Software Can Eat Shit)

Somewhere along the line, the big CAD vendors decided that:

- Drawing borders are *premium features*
- Page setup is *enterprise-only*
- Issuing a PDF requires a **four-figure annual tribute**
- And if you stop paying, your own drawings are held hostage like a bad ransom movie

All so you can:
- Put a title block on a picture
- Add a drawing number
- Add a revision
- Export a PDF

This tool exists because that entire situation is **fucking absurd**.

If all you need is a:
- Border
- Title block
- Consistent layout
- Professional-looking output

Then you do **not** need:
- A 20GB install
- A licence daemon running in the background
- A yearly invoice that makes your eyes water
- Or a EULA written by lawyers who actively despise you

---

## What It Actually Does (Before Anyone Gets Confused)

- Takes PNGs, JPGs, and PDFs
- Splits PDFs into individual sheets
- Places each sheet into a proper viewport
- Adds an ISO-style title block
- Outputs **one combined PDF drawing pack**
- Defaults to A3 landscape like a grown-up tool should

Per sheet:
- Drawing title
- Comments (that actually wrap, shockingly)

Global:
- Project
- Client
- Drawing number
- Revision
- Date
- Issuer company
- Optional logo (because branding apparently matters to people)

No modelling.  
No constraints.  
No parametric wizardry.  
Just **getting shit issued**.

---

## Who This Is For

You, when:

- The drawing was done in PowerPoint at 2am
- The “proper” CAD export looked like arse
- Someone said “can you just issue that?”
- You need it to look professional *now*
- And reinstalling SolidWorks for one PDF would make you scream

If your instinctive response to this tool is:
“Why don’t you just use AutoCAD?”

Congratulations.  
You are the reason this exists.

---

## What This Is NOT (Read This Slowly)

- Not CAD
- Not BIM
- Not parametric
- Not scale-authoritative
- Not pretending to replace anything heavy-duty

This is a **finishing tool**, not a lifestyle choice.

It does not care about your feature tree.
It does not care about mates.
It does not care about Dassault’s feelings.

---

## Technical Reality (For The Curious)

- Written in Python
- GUI via PyQt5
- PDFs via ReportLab
- PDF handling via PyMuPDF
- Runs entirely locally
- Works without internet
- Keeps working even if a multinational decides to “sunset” something

Astonishing, really.

---

## Philosophy (The Bit CAD Vendors Hate)

- If it’s your drawing, you should be allowed to print it
- A rectangle and some text should not cost thousands
- Software should shut the fuck up and do the job
- Tools should not expire because accounting said so

No cloud.  
No accounts.  
No “trial expired”.  
No watermark threatening legal action.

---

## Status

Used on real projects.  
Trusted more than licence servers.  
Maintained out of pure spite.

If it breaks, it’s because *I* broke it — not because a vendor flipped a switch.

---

## Licence

MIT.

Do whatever you want.
Fork it.
Ship it.
Rename it.
Sell it.

Just don’t add a login screen or I will personally judge you.

---

## Final Thought

This tool exists to screw over exactly one idea:

That you should need **AutoCAD or SolidWorks** just to wrap a picture in a border and press “Export PDF”.

They can absolutely get fucked.

Enjoy.
