# Usage #

- Get-VssWriters

The above command will pull the vss writer information for the local computer.

- Get-VssWriters -ComputerName svr11, svr2, $env:COMPUTER

The abovecommand will pull the vss writer information for the computers: svr11, svr2, and the localcomputer.

---

Special thanks to Sam Boutros for the initial script.
https://superwidgets.wordpress.com/category/powershell/