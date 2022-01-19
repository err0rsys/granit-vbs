# granit_vbs
An example of starting the Granit system functionality with the help of vbscript

The granit.vbs script allows to log into the Granit system on the selected database and execute a sample method from the ScriptCom component. The method displays the Script Pascal language methods supported in the given version of ScriptCOM.

Login to the Granit system is carried out using the LoginClient method, which allows, among other things, non-interactive registration of the user in the system and the assignment of an access code (AccessCode), which is used at a later stage for authentication purposes.

In case of failure during the login process the user is informed about the possible cause of the problem (in the script, for simplicity, not all possible error codes that the login method may return are handled).

When a successful login is performed, the AssignAccessCode method of the component's target interface is run, followed by the selected method that performs the business function.
