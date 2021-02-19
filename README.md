# ADCS-PKI-Scripts
AD Certificate Services and PKI scripts of interest

**Large Database reporting demo files**

You can use LargeCollector.cmd with Process-Certutil-C.ps1 to produce a CSV file of certificates from an ADCS database.

On the CA, run

`LargeCollector.cmd .\Active.log Active`

to dump active certificates with pre-selected fields in long-text form to active.log (run LargeCollector on its own to see the pre-canned query types), and then

`.\process-certutil-c.ps1 .\active.log -exportfile Active.csv`

to (fairly laboriously) convert that text log into a CSV file, useful for Excel pivot tables and PowerShell and stuff.
