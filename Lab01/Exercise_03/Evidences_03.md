# Microsoft Ms-600 (Adrián Arenilla Seco) - LAB 01


## Exercise 3: Implementing application that supports B2B
### [Go to exercise 03 instructions -->](04-Exercise-3-Implementing-application-that-supports-B2B.md)


Register a single-tenant Azure AD application.
![](Evidences/Image04a.png)


Add a Redirect URI of web applications. 

Add a Redirect URI. 

Select both Access tokens and ID tokens and save.
![](Evidences/Image04b.png)


After update the files build the multiple organization web app.
![](Evidences/Image04c.png)


Test the app.
![](Evidences/Image04d.png)


Notice Azure AD will reject the user's sign in, explaining that the user's account doesn't exist in the current tenant.
![](Evidences/Image04e.png)


Invite a guest user from another organization.
![](Evidences/Image04f.png)


Test the app.
![](Evidences/Image04g.png)


### [<-- Back to readme](../../../../)