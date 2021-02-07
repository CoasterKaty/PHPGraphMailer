<?php
/*
 *   Mail script - Katy Nicholson, April 2020
 *
 *   Retrieves and sends messages from an Exchange Online mailbox using MS Graph API
 *
 *   https://katystech.blog/2020/08/php-graph-mailer/
 */
class graphMailer {

    var $tenantID;
    var $clientID;
    var $clientSecret;
    var $Token;
    var $baseURL;

    function __construct($sTenantID, $sClientID, $sClientSecret) {
            $this->tenantID = $sTenantID;
            $this->clientID = $sClientID;
            $this->clientSecret = $sClientSecret;
        $this->baseURL = 'https://graph.microsoft.com/v1.0/';
        $this->Token = $this->getToken();
    }

    function getToken() {
            $oauthRequest = 'client_id=' . $this->clientID . '&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret=' . $this->clientSecret . '&grant_type=client_credentials';
        $reply = $this->sendPostRequest('https://login.microsoftonline.com/' . $this->tenantID . '/oauth2/v2.0/token', $oauthRequest);
            $reply = json_decode($reply['data']);
            return $reply->access_token;

    }

    function getMessages($mailbox) {
        if (!$this->Token) {
            throw new Exception('No token defined');
        }
        $messageList = json_decode($this->sendGetRequest($this->baseURL . 'users/' . $mailbox . '/mailFolders/Inbox/Messages'));
        if ($messageList->error) {
            throw new Exception($messageList->error->code . ' ' . $messageList->error->message);
        }
        $messageArray = array();

        foreach ($messageList->value as $mailItem) {
            $attachments = (json_decode($this->sendGetRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $mailItem->id . '/attachments')))->value;
            if (count($attachments) < 1) unset($attachments);
            foreach ($attachments as $attachment) {
                if ($attachment->{'@odata.type'} == '#microsoft.graph.referenceAttachment') {
                    $attachment->contentBytes = base64_encode('This is a link to a SharePoint online file, not yet supported');
                    $attachment->isInline = 0;
                }
            }
            $messageArray[] = array('id' => $mailItem->id,
                        'sentDateTime' => $mailItem->sentDateTime,
                        'subject' => $mailItem->subject,
                        'bodyPreview' => $mailItem->bodyPreview,
                        'importance' => $mailItem->importance,
                        'conversationId' => $mailItem->conversationId,
                        'isRead' => $mailItem->isRead,
                        'body' => $mailItem->body,
                        'sender' => $mailItem->sender,
                        'toRecipients' => $mailItem->toRecipients,
                        'ccRecipients' => $mailItem->ccRecipients,
                        'toRecipientsBasic' => $this->basicAddress($mailItem->toRecipients),
                        'ccRecipientsBasic' => $this->basicAddress($mailItem->ccRecipients),
                        'replyTo' => $mailItem->replyTo,
                        'attachments' => $attachments);

        }
        return $messageArray;
    }

    function deleteEmail($mailbox, $id, $moveToDeletedItems = true) {
        switch ($moveToDeletedItems) {
            case true:
                $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $id . '/move', '{ "destinationId": "deleteditems" }', array('Content-type: application/json'));
                break;
            case false:
                $this->sendDeleteRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $id);
                break;
        }
    }

    function sendMail($mailbox, $messageArgs ) {
        if (!$this->Token) {
            throw new Exception('No token defined');
        }

        /*
        $messageArgs[   subject,
                replyTo{'name', 'address'},
                toRecipients[]{'name', 'address'},
                ccRecipients[]{'name', 'address'},
                importance,
                conversationId,
                body,
                images[],
                attachments[]
                ]

        */

        foreach ($messageArgs['toRecipients'] as $recipient) {
            if ($recipient['name']) {
                $messageArray['toRecipients'][] = array('emailAddress' => array('name' => $recipient['name'], 'address' => $recipient['address']));
            } else {
                $messageArray['toRecipients'][] = array('emailAddress' => array('address' => $recipient['address']));
            }
        }
        foreach ($messageArgs['ccRecipients'] as $recipient) {
            if ($recipient['name']) {
                $messageArray['ccRecipients'][] = array('emailAddress' => array('name' => $recipient['name'], 'address' => $recipient['address']));
            } else {
                $messageArray['ccRecipients'][] = array('emailAddress' => array('address' => $recipient['address']));
            }
        }
        $messageArray['subject'] = $messageArgs['subject'];
        $messageArray['importance'] = ($messageArgs['importance'] ? $messageArgs['importance'] : 'normal');
        if (isset($messageArgs['replyTo'])) $messageArray['replyTo'] = array(array('emailAddress' => array('name' => $messageArgs['replyTo']['name'], 'address' => $messageArgs['replyTo']['address'])));
        $messageArray['body'] = array('contentType' => 'HTML', 'content' => $messageArgs['body']);
        $messageJSON = json_encode($messageArray);
        $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages', $messageJSON, array('Content-type: application/json'));

        $response = json_decode($response['data']);
        $messageID = $response->id;

        foreach ($messageArgs['images'] as $image) {
            $messageJSON = json_encode(array('@odata.type' => '#microsoft.graph.fileAttachment', 'name' => $image['Name'], 'contentBytes' => base64_encode($image['Content']), 'contentType' => $image['ContentType'], 'isInline' => true, 'contentId' => $image['ContentID']));
            $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $messageID . '/attachments', $messageJSON, array('Content-type: application/json'));
        }

        foreach ($messageArgs['attachments'] as $attachment) {
            $messageJSON = json_encode(array('@odata.type' => '#microsoft.graph.fileAttachment', 'name' => $attachment['Name'], 'contentBytes' => base64_encode($attachment['Content']), 'contentType' => $attachment['ContentType'], 'isInline' => false));
            $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $messageID . '/attachments', $messageJSON, array('Content-type: application/json'));
        }
        //Send
        $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $messageID . '/send', '', array('Content-Length: 0'));
        if ($response['code'] == '202') return true;
        return false;

    }

    function basicAddress($addresses) {
        foreach ($addresses as $address) {
            $ret[] = $address->emailAddress->address;
        }
        return $ret;
    }

    function sendDeleteRequest($URL) {
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'DELETE');
        curl_setopt($ch, CURLOPT_HTTPHEADER, array('Authorization: Bearer ' . $this->Token, 'Content-Type: application/json'));
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        curl_close($ch);
        echo $response;
    }

    function sendPostRequest($URL, $Fields, $Headers = false) {
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_POST, 1);
        if ($Fields) curl_setopt($ch, CURLOPT_POSTFIELDS, $Fields);
        if ($Headers) {
            $Headers[] = 'Authorization: Bearer ' . $this->Token;
            curl_setopt($ch, CURLOPT_HTTPHEADER, $Headers);
        }
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        $responseCode = curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
        curl_close($ch);
        return array('code' => $responseCode, 'data' => $response);
    }

    function sendGetRequest($URL) {
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_HTTPHEADER, array('Authorization: Bearer ' . $this->Token, 'Content-Type: application/json'));
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        curl_close($ch);
        return $response;
    }
}

?>
