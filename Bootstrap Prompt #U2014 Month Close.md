# Starting a Month Close (v3, Sept 2026)

Open a new chat and type:
**Start the chat from the Claude desktop app on the MacBook, inside the Weekly and Monthly Reports project.** That is what gives the chat the workspace folder and the computer link the close needs. A chat started from the web app, the phone, or outside the project will report it can't reach your Mac.


> **Let's start [Month Year] close.**

Attach whatever you have (food statements first is fine; the rest can follow in the same chat):
- 4 Square Sales Reports (Studio, 1899, NMWA, DT)
- IGF Statement + Baldor Usage Reports
- ADP PDFs (one per pay date)

That's it. The saved skill `slipstream-month-close` reads the Playbook and memory, checks the sheets, derives pay dates from the ADP PDFs, applies the standing rules (NMWA fee, "Other" tender, Linen/Catering defaults, 1899 mgmt-fee scan), and only asks for what it can't derive — usually nothing, occasionally the Commissary food cost for a NMWA catering event.

If a fresh chat seems lost: "use the slipstream-month-close skill."
