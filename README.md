##Outlook Rules Export as plain text##
A rewrite to write rules as plain human-readable text


##Supported Rule Types##
* The tool currently only works for the following types of rules (will add support for other types of rules as i need them or as requested)
    * Condition: "From Address" Actions: "Move-To-Folder | Copy-To-Folder"
    * Condition: "Subject Contains" Actions: "Move-To-Folder | Copy-To-Folder"
    * Condition: "Body Contains" Actions: "Move-To-Folder | Copy-To-Folder"
    
##Credits##
* This is a rewrite of the https://github.com/iloveitaly/outlook-gmail-rules-migration project from ruby and vb to c#
* Thanks to iloveitaly (Michael Bianco) https://github.com/iloveitaly for his work figuring out the various ways Outlook stores its rules https://github.com/iloveitaly/outlook-gmail-rules-migration 
