# GetCMUpdateGroupDetails
Get details about the updates in a Configuration Manager Update Group.

The script gives you details about the updates in an update group including the size of each update included. Mind that a download for a certain kb article may consist of different updates for different releases (like x86, x64, amd etc.) or different languages. The report should help to figure out how much content will actually be downloaded to a certain group of machines.
The output is an html file that is by default in the same folder as the script.

Finally, mind that this code is provided as-is with no warrenties. If you find any issues or if you have ideas for improvements feel free to contact me. Thanks!
