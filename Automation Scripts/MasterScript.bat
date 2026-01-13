@echo off
echo Currently logged-in user: %USERNAME%

"start cmd /k python "C:\Users\%USERNAME%\QE GIT\Bayobab-Py-N-Dump\MailsUpdater_Automated Version.py"
start cmd /k python "C:\Users\%USERNAME%\QE GIT\Bayobab-Py-N-Dump\Fixed_Performance_Report_Updater.py"




"start cmd /k python "C:\Users\%USERNAME%\MTN Group\CIM Reports Hub - Documents\Automation_Scripts\Python Scripts\MailsUpdater_Automated Version.py""
"start cmd /k python "C:\Users\%USERNAME%\QE GIT\Bayobab-Py-N-Dump\MailsUpdater_Automated Version.py"
"start cmd /k python "C:\Users\%USERNAME%\MTN Group\CIM Reports Hub - Documents\Automation_Scripts\Python Scripts\MailsUpdater_Automated Version - Reconciler.py""
"start cmd /k python "C:\Users\%USERNAME%\MTN Group\CIM Reports Hub - Documents\Automation_Scripts\Fixed_Performance_Report_Updater.py""
