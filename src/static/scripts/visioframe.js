

window.onload = function() {
    setVisioFileInIFrame();
    //getFileData();
};

var redeemFile=function(){

}
function getAuthToken(){
    microsoftTeams.getContext(function (context) {
    });
}
function getFileData(){
    fetch(getRedeemUri(), getRequestData())
      .then( (response) => { 
         console.log(response);
      });
}
function getRedeemUri(){
    
    return "https://microsoft.sharepoint-df.com/teams/Visio/_api/v2.0/sites/root/items/94f5c607-26ea-4ff8-8c46-b6a0c9487bd8/driveItem?select=OpenWith,officeBundle,file,size,name,@microsoft.graph.downloadUrl,etag,currentUserRole,webUrl,sensitivityLabel,sharepointIds,webDavUrl&action=open&ump=1";
}
function getRequestData(){
    var request={};
    request.method="POST";
    var headers={};
    var boundary="6b778aec-cef2-4979-814c-1293917435b5";
    var authToken=getAuthToken();
    headers["Accept"]="*/*";
    headers["Content-Type"]="multipart/form-data;boundary="+boundary;
    headers["Authorization"]="Bearer "+authToken;
    request.headers=headers;
    request.body=getObjToPost(boundary,authToken);
}
function getAuthToken(){
    return "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Imh1Tjk1SXZQZmVocTM0R3pCRFoxR1hHaXJuTSIsImtpZCI6Imh1Tjk1SXZQZmVocTM0R3pCRFoxR1hHaXJuTSJ9.eyJhdWQiOiJodHRwczovL21pY3Jvc29mdC5zaGFyZXBvaW50LWRmLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0Ny8iLCJpYXQiOjE1OTU3Mzg2NjYsIm5iZiI6MTU5NTczODY2NiwiZXhwIjoxNTk1NzQyNTY2LCJhY3IiOiIxIiwiYWlvIjoiQVZRQXEvOFFBQUFBQUVLMzN3VjNLd1M2SUFKVGRpK3h2QWtwUnhwNXk5Sll4VWNucjgrbStZNEMvL1dpbDhRakRKRHJNWFFVV21yTHpFb2hLUzFCR0ZHb2JCc1pzRVFrMDdhWEREQ2xYTTFHbW9DRmNUbnJNZEk9IiwiYW1yIjpbIndpYSIsIm1mYSJdLCJhcHBfZGlzcGxheW5hbWUiOiJNaWNyb3NvZnQgVGVhbXMgV2ViIENsaWVudCIsImFwcGlkIjoiNWUzY2U2YzAtMmIxZi00Mjg1LThkNGItNzVlZTc4Nzg3MzQ2IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJQcmFzYWQiLCJnaXZlbl9uYW1lIjoiVXJ2YXNoaSIsImlkdHlwIjoidXNlciIsImluX2NvcnAiOiJ0cnVlIiwiaXBhZGRyIjoiMTgzLjgzLjE0My42MCIsIm5hbWUiOiJVcnZhc2hpIFByYXNhZCIsIm9pZCI6ImVmMTEwMDI3LTk1MWYtNDk4Zi05NTlmLTM3YzE0YWFmMWQ0ZiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMTQ2NzczMDg1LTkwMzM2MzI4NS03MTkzNDQ3MDctMjU3MzQ2MyIsInB1aWQiOiIxMDAzMjAwMDhCNzVDQjhEIiwicmgiOiIwLkFSb0F2NGo1Y3ZHR3IwR1JxeTE4MEJIYlI4RG1QRjRmSzRWQ2pVdDE3bmg0YzBZYUFGYy4iLCJzY3AiOiJNeUZpbGVzLldyaXRlIFNpdGVzLkZ1bGxDb250cm9sLkFsbCBTaXRlcy5NYW5hZ2UuQWxsIFVzZXIuUmVhZFdyaXRlLkFsbCIsInNpZCI6ImQ4Mzk0N2NjLWE5NWMtNGFiNi1hMTcxLWRkNDYwNDU2YTc0NSIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IjlZRGtmc0JTSFpOV3FJM1FUQlAtRVRZWGVnN3ZsMHM2dWRyOGZSekxPcEEiLCJ0aWQiOiI3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJ1bmlxdWVfbmFtZSI6InVycHJhc2FkQG1pY3Jvc29mdC5jb20iLCJ1cG4iOiJ1cnByYXNhZEBtaWNyb3NvZnQuY29tIiwidXRpIjoiUmtYRUtFbS1UMHVjd2FLRTNiQWZBQSIsInZlciI6IjEuMCJ9.noRY1tFwr2xeuv9tCaa_Octr79XZ3PpwIkqeLAYacwo3YOcQXqD6zPHHO8dA15utDW2hBkm_MDTTGl5wW1xu1GvtjaB9Rolobqi1AT4KaDq7HcD53agZiY5XgMUcmgUaNphuTvRiou8y2Bq1c_i8PS8JvKmQjS82CBhA7b_GhfyjM9_5ljBKEuTdyiEOBcY-Wor9oMqX-ES0_uu395EH_X15Tr3C7Zm6yo3q9-IFtXXWex4fjtKyZlMFt8CtIh-TDRSl-_VHrx31jCwte-fugxORIxkVZDhYCIVfohkhHwrGj4-3e0f6FC8VizLgqjCKXQoEY7MYJcut0b3_YWXgtA";
}
function setVisioFileInIFrame(){
    var iframe = document.getElementById('visioFrame');
   // iframe.src = "https://microsoft.sharepoint-df.com/teams/Visio/_layouts/15/PreAuth.aspx?sourcedoc={94f5c607-26ea-4ff8-8c46-b6a0c9487bd8}&action=view#access_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6InhYbFQwSUp4MlZCVEJFeVFMdGtFOFY2ZkhwZyJ9%2EeyJhdWQiOiJ3b3BpL21pY3Jvc29mdC5zaGFyZXBvaW50LWRmLmNvbUA3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJpc3MiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDBAOTAxNDAxMjItODUxNi0xMWUxLThlZmYtNDkzMDQ5MjQwMTliIiwibmJmIjoiMTU5NTY2OTY3OCIsImV4cCI6IjE1OTU2Njk5NzgiLCJuYW1laWQiOiIxMDAzMjAwMDhiNzVjYjhkIiwibmlpIjoidXJuOmZlZGVyYXRpb246bWljcm9zb2Z0b25saW5lIiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDA4Yjc1Y2I4ZEBsaXZlLmNvbSIsImlzdXNlciI6InRydWUiLCJhcHBpZCI6IjVlM2NlNmMwLTJiMWYtNDI4NS04ZDRiLTc1ZWU3ODc4NzM0NiIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsImlwYWRkciI6IjE4My44My4xNDMuNjAiLCJ3b3BpX2FwIjoibXlmaWxlcy53cml0ZSBhbGxzaXRlcy5mdWxsY29udHJvbCBhbGxzaXRlcy5tYW5hZ2UgYWxscHJvZmlsZXMud3JpdGUiLCJ3b3BpX3R0IjoiUHJlQXV0aFRva2VuIiwiYXBwY3R4IjoiOTRmNWM2MDcyNmVhNGZmODhjNDZiNmEwYzk0ODdiZDg7U21HenJGaWJHaXlGd1IzM0FERURJVm5FZmFvPTtEZWZhdWx0OzsxQjAzQzQzMUFFRjtUcnVlOzs1ZTNjZTZjMC0yYjFmLTQyODUtOGQ0Yi03NWVlNzg3ODczNDY7MjYyODtmZTc3Njk5Zi0yMDJmLTAwMDAtY2U1Ny0yYjQ3NDdhYjE1ZmIifQ%2EeXiR0j3Ok9AUZMzYC7mySgGvPjtO73lPVp%5FLcQl5cKfXiNaFMn1N3SnIfwsOulgggV0RldUH7Mf2iiju6nsglu%2D%5Fm1BifF44hzgDH6AMAopq4g1X9tv3%2DyDJ2vUafIwzSgdgXIJZUDKzHHknn%5FYjpZNr7PehpSaeyiUHS8pFX4hupBpyNLgq4Yqw8VnpAGflj7OiPbaTjILjuU%5F6HyLi5jhRYt7aPj4Go4zLaoip%5F8FT9qBstGoNxVS9Khgk%2DUJZbNVE4AQjjmbUVI38ExsCjDHxsYlrO0UHQJtpuXkehvENxioeuxVRh%2DiQGHBQIB3gdAkrr9smh56F2vNXeakwLg";
   iframe.src="https://urprasad-ts3.fareast.corp.microsoft.com/th/FrameWAC.aspx?Fi=anonymous%7EDocument2%2Evsdx&Action=Edit&Application=Visio&transport=wopi&wachost=urprasad-ts3.fareast.corp.microsoft.com&uiembed=1";
}
function getObjToPost(boundary,authToken){
    var endOfLine="\n";
    var objToPost="--"+boundary+"\nContent-Disposition: form-data;name=data\nIf-Match: * \nAccept-Language: en-us \nprefer: redeemSharingLink,getShortLivedDownloadUrl \nContent-Type: application/json \n";
    objToPost += "Authorization: Bearer " + authToken + endOfLine;
    objToPost += "X-HTTP-Method-Override: GET";
    objToPost +=  "\n";
    objToPost += "\n{} \n";
    objToPost+="--"+boundary+"--";
    return objToPost;
}