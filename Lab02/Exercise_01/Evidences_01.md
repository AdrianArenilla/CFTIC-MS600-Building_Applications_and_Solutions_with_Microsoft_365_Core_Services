# Microsoft Ms-600 (Adrián Arenilla Seco) - LAB 02


## Exercise 1: Using query parameters when querying Microsoft Graph via HTTP
### [Go to exercise 01 instructions -->](02-Exercise-1-Using-query-parameters-when-querying-Microsoft-Graph-via-HTTP.md)


`https://graph.microsoft.com/v1.0/groups`
![](Evidences/Image02a.png)


`https://graph.microsoft.com/v1.0/groups?$select=displayName,mail,visibility`
![](Evidences/Image02b.png)


`https://graph.microsoft.com/v1.0/groups?$select=displayName,mail,visibility,id`
![](Evidences/Image02c.png)


`https://graph.microsoft.com/v1.0/groups?$select=displayName,mail,visibility,id&$orderBy=displayName`
![](Evidences/Image02d.png)


`https://graph.microsoft.com/v1.0/groups?$select=displayName,mail,visibility,id&$orderBy=displayName desc`
![](Evidences/Image02e.png)


`https://graph.microsoft.com/v1.0/me/messages`
![](Evidences/Image02f.png)


`https://graph.microsoft.com/v1.0/me/messages?$filter=hasAttachments eq true`
![](Evidences/Image02g.png)


`https://graph.microsoft.com/v1.0/me/messages?$filter=from/emailAddress/address eq 'no-reply@microsoft.com'`
![](Evidences/Image02h.png)


`https://graph.microsoft.com/v1.0/me/messages?$filter=startswith(subject,'my')`
![](Evidences/Image02i.png)


`https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c: c eq 'Unified')`
![](Evidences/Image02j.png)


`https://graph.microsoft.com/v1.0/me/events`
![](Evidences/Image02k.png)


`https://graph.microsoft.com/v1.0/me/events?$top=5`
![](Evidences/Image02l.png)


`https://graph.microsoft.com/v1.0/me/events?$top=5&$skip=5`
![](Evidences/Image02m.png)


`https://graph.microsoft.com/v1.0/me/messages?$filter=hasAttachments eq true`
![](Evidences/Image02n.png)


`https://graph.microsoft.com/v1.0/me/messages?$filter=hasAttachments eq true&$expand=attachments`
![](Evidences/Image02o.png)


`https://graph.microsoft.com/v1.0/me/messages?$filter=hasAttachments eq true&$expand=attachments($select=id,name,contentType,size)`
![](Evidences/Image02p.png)


`https://graph.microsoft.com/v1.0/me/messages?$count=true`
![](Evidences/Image02q.png)


`https://graph.microsoft.com/v1.0/me/messages?$count=true&$filter=isRead eq false`
![](Evidences/Image02r.png)


`https://graph.microsoft.com/v1.0/me/messages?$search="business"`
![](Evidences/Image02s.png)


`https://graph.microsoft.com/v1.0/me/messages?$search="subject:business"`
![](Evidences/Image02t.png)


### [<-- Back to readme](../../../../)