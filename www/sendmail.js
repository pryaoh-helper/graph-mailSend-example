

const graphUrl = 'https://graph.microsoft.com/v1.0/';
$('#mailForm').submit(function(e) {
    e.preventDefault();    
    let accessToken = $('#tokenParagraph').html();

    console.log('accessToken: ' + accessToken);
    const sourceEmail = $('#smail').val();
    const targetEmail = $('#tmail').val();


    const messageBody = {
        "message": {
          "subject": "EP Mail Test 입니다.",
          "body": {
            "contentType": "Text",
            "content": "EP Mail Test 입니다. ~~~"
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": targetEmail
              }
            }
          ],
          "ccRecipients": [
            
          ]
        },
        "saveToSentItems": "false"
      }

    $.ajax({
        type: 'POST',
        url: `${graphUrl}/users/${sourceEmail}/sendMail`,
        contentType: "application/json;odata=verbose",
        headers: {                     
            'Authorization': `Bearer ${accessToken}`
        },
        
        data: JSON.stringify(messageBody)
    })
    .done(function(data) {
        console.log(data);
    })
    .fail(function(err){
        console.warn(err);
    });
});