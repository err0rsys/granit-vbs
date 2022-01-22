# granit-vbs
An example of running the Granit system functionality using a vbscript.

The granit.vbs script allows you to remotely log into the Granit system on the selected database (eg GRANIT) and execute a sample method from the ScriptCom COM+ component. The method displays the Script Pascal language (Pascal fork) methods supported in a given version of ScriptCOM.

Logging into the Granit system is done using the LoginClient method, which allows you to, among other things, non-interactively register a user into the system and assign an access code (AccessCode) that is used later for authentication purposes.

In case of failure during the login process, the user is informed about the possible cause of the problem (in the script, for simplicity, all possible error codes that the login method may return are not supported).

After a successful login, the AssignAccessCode method of the component's target interface is run, followed by the selected method that implements the business function.
