<?php
/*
 *   Mail script - Katy Nicholson, April 2020
 *
 *   Retrieves and sends messages from an Exchange Online mailbox using MS Graph API
 *
 *   https://katystech.blog/2020/08/php-graph-mailer/
 */

class graphMailerException extends Exception { }

class graphMailer {

    private $tenantID;
    private $clientID;
    private $clientSecret;
    private $baseURL;
    private $Token;

    function __construct($sTenantID,$sClientID,$sClientSecret) {
        $this->tenantID     = $sTenantID;
        $this->clientID     = $sClientID;
        $this->clientSecret = $sClientSecret;
        $this->baseURL      = 'https://graph.microsoft.com/v1.0/';
        $this->Token        = $this->getToken();
    }

    function getToken() {
        $oauthRequest = http_build_query([
          'client_id'   => $this->clientID,
          'scope'       => 'https://graph.microsoft.com/.default&client_secret='.$this->clientSecret,
          'grant_type'  => 'client_credentials',
        ],NULL,'&amp;');
        $reply = $this->sendPostRequest('https://login.microsoftonline.com/'.$this->tenantID.'/oauth2/v2.0/token',$oauthRequest);
        $reply = json_decode($reply['data']);
        return $reply->access_token;
    }

    function getMessages($mailbox) {
        if (!$this->Token) { throw new graphMailerException('No token defined'); }
        $messageList = json_decode($this->sendGetRequest($this->baseURL.'users/'.$mailbox.'/mailFolders/Inbox/Messages'));
        if ($messageList->error) {
            throw new graphMailerException($messageList->error->code.' '.$messageList->error->message);
        }
        $messageArray = [];

        foreach ($messageList->value as $mailItem) {
            $attachments = (json_decode($this->sendGetRequest($this->baseURL.'users/'.$mailbox.'/messages/'.$mailItem->id.'/attachments')))->value;
            if (count($attachments) < 1) { $attachments = []; }
            foreach ($attachments as $attachment) {
                if ($attachment->{'@odata.type'} == '#microsoft.graph.referenceAttachment') {
                    $attachment->contentBytes = base64_encode('This is a link to a SharePoint online file, not yet supported');
                    $attachment->isInline = 0;
                }
            }
            $messageArray[] = [
              'id'                  => $mailItem->id,
              'sentDateTime'        => $mailItem->sentDateTime,
              'subject'             => $mailItem->subject,
              'bodyPreview'         => $mailItem->bodyPreview,
              'importance'          => $mailItem->importance,
              'conversationId'      => $mailItem->conversationId,
              'isRead'              => $mailItem->isRead,
              'body'                => $mailItem->body,
              'sender'              => $mailItem->sender,
              'toRecipients'        => $mailItem->toRecipients,
              'ccRecipients'        => $mailItem->ccRecipients,
              'toRecipientsBasic'   => $this->basicAddress($mailItem->toRecipients),
              'ccRecipientsBasic'   => $this->basicAddress($mailItem->ccRecipients),
              'replyTo'             => $mailItem->replyTo,
              'attachments'         => $attachments,
            ];

        }
        return $messageArray;
    }

    function deleteEmail($mailbox,$id,$moveToDeletedItems = TRUE) {
        if (!$moveToDeletedItems) { $this->sendDeleteRequest($this->baseURL.'users/'.$mailbox.'/messages/'.$id); } else {
            $this->sendPostRequest(
              $this->baseURL.'users/'.$mailbox.'/messages/'.$id.'/move',
              json_encode(['destinationId' => 'deleteditems']),
              ['Content-type: application/json']
            );
        }
    }

    function sendMail($mailbox,$messageArgs) {
        if (!$this->Token) { throw new graphMailerException('No token defined'); }

        /*
        $mailArgs =  [
            'subject'       => 'Test message',
            'replyTo'       => ['address' => 'address@email.com','name' => 'Katy'],
            'toRecipients'  => [
                ['address'  => 'address@email.com',  'name' => 'Neil'],         // Name is optional
                ['address'  => 'address2@email.com', 'name' => 'Someone'],
             ],
            'ccRecipients'  => [
                ['address'      => 'address@email.com', 'name' => 'Neil'],      // Name is optional
                ['address'      => 'address2@email.com','name' => 'Someone'],
            ],
            'importance'     => 'normal',
            'conversationId' => '', // Optional, use if replying to an existing email to keep them chained properly in outlook
            'body'           => '<html>Blah blah blah</html>',
            'images'         => [ // Array of arrays so you can have multiple images. These are inline images. Everything else in attachments.
                ['Name' => 'blah.jpg','ContentType' => 'image/jpeg','Content' => 'results of file_get_contents(blah.jpg)','ContentID' => 'cid:blah'],
            ],
            'attachments'    => [
                ['Name' => 'blah.pdf', 'ContentType' => 'application/pdf', 'Content' => 'results of file_get_contents(blah.pdf)'],
            ]
        ];

        $graphMailer = new graphMailer($sTenantID,$sClientID,$sClientSecret);
        $graphMailer->sendMail('helpdesk@contoso.com',$mailArgs);
        unset($graphMailer);
        */

        $messageArray = [];
        foreach(['toRecipients','ccRecipients'] as $arr) {
            foreach($messageArgs[$arr] as $recipient) {
                $add = ['emailAddress' => ['address' => $recipient['address']]];
                if ($recipient['name']) { $add['emailAddress']['name'] = $recipient['name']; }
                $messageArray[$arr][] = $add;
            }
        }

        $messageArray['subject']    = $messageArgs['subject'];
        $messageArray['importance'] = ($messageArgs['importance'] ? $messageArgs['importance'] : 'normal');
        if (isset($messageArgs['replyTo'])) {
            $messageArray['replyTo'] = [['emailAddress' => ['address' => $messageArgs['replyTo']['address'],'name' => $messageArgs['replyTo']['name']]]];
        }
        $messageArray['body']       = ['contentType' => 'HTML','content' => $messageArgs['body']];
        $response                   = $this->sendPostRequest(
          $this->baseURL.'users/'.$mailbox.'/messages',
          json_encode($messageArray),
          ['Content-type: application/json']
        );

        $response   = json_decode($response['data']);
        $messageID  = $response->id;

        foreach ($messageArgs['images'] as $image) {
            $messageJSON = json_encode([
              '@odata.type'     => '#microsoft.graph.fileAttachment',
              'name'            => $image['Name'],
              'contentBytes'    => base64_encode($image['Content']),
              'contentType'     => $image['ContentType'],
              'isInline'        => TRUE,
              'contentId'       => $image['ContentID'],
            ]);
            $this->sendPostRequest($this->baseURL.'users/'.$mailbox.'/messages/'.$messageID.'/attachments',$messageJSON,['Content-type: application/json']);
        }

        foreach ($messageArgs['attachments'] as $attachment) {
            $messageJSON = json_encode([
              '@odata.type'     => '#microsoft.graph.fileAttachment',
              'name'            => $attachment['Name'],
              'contentBytes'    => base64_encode($attachment['Content']),
              'contentType'     => $attachment['ContentType'],
              'isInline'        => FALSE,
            ]);
            $this->sendPostRequest($this->baseURL.'users/'.$mailbox.'/messages/'.$messageID.'/attachments',$messageJSON,['Content-type: application/json']);
        }

        //Send
        $response = $this->sendPostRequest($this->baseURL.'users/'.$mailbox.'/messages/'.$messageID.'/send','',['Content-Length: 0']);
        return ((int) $response['code'] == 202);
    }

    function basicAddress($addresses) {
        $ret = [];
        foreach ($addresses as $address) { $ret[] = $address->emailAddress->address; }
        return $ret;
    }

    function sendDeleteRequest($URL) {
        $ch = curl_init($URL);
        curl_setopt_array($ch,[
            CURLOPT_CUSTOMREQUEST   => 'DELETE',
            CURLOPT_HTTPHEADER      => ['Authorization: Bearer '.$this->Token,'Content-Type: application/json'],
            CURLOPT_RETURNTRANSFER  => TRUE,
        ]);
        echo curl_exec($ch);
        curl_close($ch);
    }

    function sendPostRequest($URL,$Fields,$Headers = FALSE) {
        $ch = curl_init($URL);
        $opts = [
            CURLOPT_POST            => TRUE,
            CURLOPT_RETURNTRANSFER  => TRUE,
        ];
        if ($Fields)    { $opts[CURLOPT_POSTFIELDS] = $Fields; }
        if ($Headers)   {
            $Headers[] = 'Authorization: Bearer '.$this->Token;
            $opts[CURLOPT_HTTPHEADER] = $Headers;
        }
        curl_setopt_array($ch,$opts);

        $response = curl_exec($ch);
        $responseCode = curl_getinfo($ch,CURLINFO_RESPONSE_CODE);
        curl_close($ch);
        return ['code' => $responseCode,'data' => $response];
    }

    function sendGetRequest($URL) {
        $ch = curl_init($URL);
        curl_setopt_array($ch,[
            CURLOPT_HTTPHEADER      => ['Authorization: Bearer '.$this->Token,'Content-Type: application/json'],
            CURLOPT_RETURNTRANSFER  => TRUE,
        ]);
        $response = curl_exec($ch);
        curl_close($ch);
        return $response;
    }
}

?>