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

        if(isset($reply->error)) {
            throw new Exception($reply->error_description);
        }

        return $reply->access_token;
    }

	function createMessageJSON($messageArgs, $addMessageEnvelope = False) {
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
		$messageArray = array();
		if (array_key_exists('toRecipients', $messageArgs)) {
			foreach ($messageArgs['toRecipients'] as $recipient) {
				if (array_key_exists('name', $recipient)) {
					$messageArray['toRecipients'][] = array('emailAddress' => array('name' => $recipient['name'], 'address' => $recipient['address']));
				} else {
					$messageArray['toRecipients'][] = array('emailAddress' => array('address' => $recipient['address']));
				}
			}
		}
		if (array_key_exists('ccRecipients', $messageArgs)) {
			foreach ($messageArgs['ccRecipients'] as $recipient) {
				if (array_key_exists('name', $recipient)) {
					$messageArray['ccRecipients'][] = array('emailAddress' => array('name' => $recipient['name'], 'address' => $recipient['address']));
				} else {
					$messageArray['ccRecipients'][] = array('emailAddress' => array('address' => $recipient['address']));
				}
			}
		}
		if (array_key_exists('bccRecipients', $messageArgs)) {
			foreach ($messageArgs['bccRecipients'] as $recipient) {
				if (array_key_exists('name', $recipient)) {
					$messageArray['bccRecipients'][] = array('emailAddress' => array('name' => $recipient['name'], 'address' => $recipient['address']));
				} else {
					$messageArray['bccRecipients'][] = array('emailAddress' => array('address' => $recipient['address']));
				}
			}
		}
		if (array_key_exists('subject', $messageArgs)) $messageArray['subject'] = $messageArgs['subject'];
		if (array_key_exists('importance', $messageArgs)) $messageArray['importance'] = $messageArgs['importance'];
        if (isset($messageArgs['replyTo'])) $messageArray['replyTo'] = array(array('emailAddress' => array('name' => $messageArgs['replyTo']['name'], 'address' => $messageArgs['replyTo']['address'])));
        if (array_key_exists('body', $messageArgs)) $messageArray['body'] = array('contentType' => 'HTML', 'content' => $messageArgs['body']);
		if ($addMessageEnvelope) {
			$messageArray = array('message'=>$messageArray);
			if (array_key_exists('comment', $messageArgs)) {
				$messageArray['comment'] = $messageArgs['comment'];
			}
			if (count($messageArray['message']) == 0) unset($messageArray['message']);
		}
        return json_encode($messageArray);
	}

	function getMessage($mailbox, $id, $folder = "") {
		if ($folder != "") {
			$response = $this->sendGetRequest($this->baseURL . 'users/' . $mailbox . '/mailFolders/'.$folder.'/messages/' . $id);
		} else {
			$response = $this->sendGetRequest($this->baseURL . 'users/' . $mailbox . '/Messages/'.$id);
		}
		$message = json_decode($response);
		if (!array_key_exists("error", $message)) {
			return $message;
		}
		return False;
	}

    function getMessages($mailbox, $folder = "inbox", $filter = "", $getAttachments = True) {
		$messageBlankArray = array(
				'sentDateTime'=>"",
				'subject'=>"",
				'bodyPreview'=>"",
                'importance'=>"",
                'conversationId'=>"",
				'isRead'=>"",
                'body'=>"",
				'sender'=>"",
                'toRecipients'=>array(),
                'ccRecipients'=>array(),
				'replyTo'=>array(),
                'images'=>array(),
                'attachments'=>array(),
        );

        if (!$this->Token) {
            throw new Exception('No token defined');
        }
		$filter .= '&$expand=singleValueExtendedProperties($filter=id eq \'string 0x0070\')'; // Get ConversationTopic
		if ($filter != "") $filter = str_replace(":", "%3A", $filter);
		if ($filter != "") $filter = str_replace("/", "%2F", $filter);
		if ($filter != "") $filter = "?".str_replace(" ", "%20", $filter);
		$url = $this->baseURL . 'users/' . $mailbox . '/mailFolders/'.$folder.'/messages'.$filter;
        $messageList = json_decode($this->sendGetRequest($url));

        if (isset($messageList->error)) {
            throw new Exception($messageList->error->code . ' ' . $messageList->error->message);
        }
        $messageArray = array();
        foreach ($messageList->value as $mailItem) {
			$messageBlank = json_decode(json_encode($messageBlankArray), FALSE);
			foreach($mailItem as $property => $value) {
				$messageBlank->$property = $value;
			}
			$mailItem = $messageBlank;
			if ($getAttachments) {
				$attachments = (json_decode($this->sendGetRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $mailItem->id . '/attachments')))->value;
				if (count($attachments) < 1) {
					unset($attachments);
				} else {
					foreach ($attachments as $attachment) {
						if ($attachment->{'@odata.type'} == '#microsoft.graph.referenceAttachment') {
							$attachment->contentBytes = base64_encode('This is a link to a SharePoint online file, not yet supported');
							$attachment->isInline = 0;
						}
					}
				}
			}

			foreach ($mailItem->singleValueExtendedProperties as $singleValueExtendedProperty) {
				if ($singleValueExtendedProperty->id == "String 0x70") {
					$ConversationTopic = $singleValueExtendedProperty->value;
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
                        'attachments' => isset($attachments) ? $attachments : null,
						'conversationTopic' => $ConversationTopic
					);

        }
        return $messageArray;
    }

	function getFolderId($mailbox, $folderName) {
		$response = $this->sendGetRequest($this->baseURL .'users/'.$mailbox.'/mailFolders?$select=displayName&$top=100');
		$folderList = json_decode($response)->value;
		foreach ($folderList as $folder) {
            //echo $folder->displayName.PHP_EOL;
			if ($folder->displayName == $folderName) {
				return $folder->id;
			}
		}
		// Now try subfolders
		foreach ($folderList as $folder) {
            //echo $folder->displayName.PHP_EOL;
			$response = $this->sendGetRequest($this->baseURL .'users/'.$mailbox.'/mailFolders/'.$folder->id.'/childFolders?$select=displayName&$top=100');
			$childFolderList = json_decode($response)->value;
			foreach ($childFolderList as $childFolder) {
//echo $childFolder->displayName.PHP_EOL;
				if ($childFolder->displayName == $folderName) {
					return $childFolder->id;
				}
			}
		}
		return false;
	}

    function deleteEmail($mailbox, $id, $moveToDeletedItems = true, $folder = "") {
        switch ($moveToDeletedItems) {
            case true:
				if ($folder != "") {
					$response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/mailFolders/'.$folder.'/messages/' . $id . '/move', '{ "destinationId": "deleteditems" }', array('Content-type: application/json'));
				} else {
					$response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $id . '/move', '{ "destinationId": "deleteditems" }', array('Content-type: application/json'));
				}
                break;
            case false:
				if ($folder != "") {
					$response = $this->sendDeleteRequest($this->baseURL . 'users/' . $mailbox . '/mailFolders/'.$folder.'/messages/' . $id);
				} else {
					//$response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $id . '/move', '{ "destinationId": "recoverableitemspurges" }', array('Content-type: application/json'));
					$response = $this->sendDeleteRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $id);
				}
                break;
        }
		$responseObj = json_decode($response);
		if ($responseObj) {
			if ($responseObj->error->code == "ErrorItemNotFound") {
				return $responseObj;
			} else {
				return False;
			}
		}
		return True;
    }

    function sendMail($mailbox, $messageArgs, $deleteAfterSend = null ) {
        if (!$this->Token) {
            throw new Exception('No token defined');
        }

        if($deleteAfterSend === null) {
            throw new Exception('Argument deleteAfterSend (true|false) is required');
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
        $messageJSON = $this->createMessageJSON($messageArgs);
        $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages', $messageJSON, array('Content-type: application/json'));
//var_dump($response);
        $responsedata = json_decode($response['data']);
        $messageID = $responsedata->id;

        if(isset($messageArgs['images'])) {
            foreach ($messageArgs['images'] as $image) {
                $messageJSON = json_encode(array('@odata.type' => '#microsoft.graph.fileAttachment', 'name' => $image['Name'], 'contentBytes' => base64_encode($image['Content']), 'contentType' => $image['ContentType'], 'isInline' => true, 'contentId' => $image['ContentID']));
                $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $messageID . '/attachments', $messageJSON, array('Content-type: application/json'));
            }
        }

        if(isset($messageArgs['attachments'])) {
            foreach ($messageArgs['attachments'] as $attachment) {
                $messageJSON = json_encode(array('@odata.type' => '#microsoft.graph.fileAttachment', 'name' => $attachment['Name'], 'contentBytes' => base64_encode($attachment['Content']), 'contentType' => $attachment['ContentType'], 'isInline' => false));
                $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $messageID . '/attachments', $messageJSON, array('Content-type: application/json'));
            }
        }
        //Send
        $response = $this->sendPostRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $messageID . '/send', '', array('Content-Length: 0'));

		if ($deleteAfterSend) {
			//Delete the message if $deleteAfterSend
			$this->deleteEmail($mailbox, $messageID, False, "sentitems");
			$messageID = true;
		}

        if ($response['code'] == '202') return $messageID;
        return false;
    }

	function reply($mailbox, $id, $messageArgs, $saveOnly, $deleteAfterSend) {
		$messageJSON = $this->createMessageJSON($messageArgs, True);
		$response = $this->sendPostRequest($this->baseURL . 'users/'.$mailbox.'/messages/'.$id.'/createReply', $messageJSON, array('Content-type: application/json'));
		$responsedata = json_decode($response['data']);
		$messageId = $responsedata->id;
		if (!$saveOnly) {
			if ($deleteAfterSend) sleep(3); // Sleep for 3 seconds before sending this message to try and flush the mailbox
			$response = $this->sendPostRequest($this->baseURL . 'users/'.$mailbox.'/messages/'.$messageId.'/send', false, array('Content-type: application/json'));
			if ($deleteAfterSend) {
				//Delete the message if $deleteAfterSend
				// We do not know how long it will take to get to the sentitems folder so there is no guaranteee
				$messageId = $this->deleteEmail($mailbox, $messageId, False, "sentitems"); // True if deleted False if not
			}
		}
		return $messageId;
	}

	function updateMessage($mailbox, $id, $Fields, $folder = "") {
		if ($folder != "") {
			$response = $this->sendPatchRequest($this->baseURL . 'users/' . $mailbox . '/mailFolders/'.$folder.'/messages/' . $id , json_encode($Fields), array('Content-type: application/json'));
		} else {
			$response = $this->sendPatchRequest($this->baseURL . 'users/' . $mailbox . '/messages/' . $id, json_encode($Fields), array('Content-type: application/json'));
		}
		if ($response['code'] == 200) {
			return True;
		} else {
			return $response;
		}
	}

	function createEvent($mailbox, $eventArray) {
		$eventJSON = json_encode($eventArray);
		$response = $this->sendPostRequest($this->baseURL . 'users/'.$mailbox.'/calendar/events', $eventJSON, array('Content-type: application/json'));
		if ($response['code'] == 201) {
			$responsedata = json_decode($response['data']);
			return $responsedata->id;
		} else {
			return $response;
		}
	}

	function updateEvent($mailbox, $id, $eventArray) {
		$eventJSON = json_encode($eventArray);
		$response = $this->sendPatchRequest($this->baseURL . 'users/'.$mailbox.'/calendar/events/'.$id, $eventJSON, array('Content-type: application/json'));
		if ($response['code'] == 200) {
			return True;
		} else {
			return $response;
		}
	}

	function cancelEvent($mailbox, $id, $comment) {
		$eventJSON = json_encode(array('comment'=>$comment));
		$response = $this->sendPostRequest($this->baseURL . 'users/'.$mailbox.'/calendar/events/'.$id."/cancel", $eventJSON, array('Content-type: application/json'));
		if ($response['code'] == 202) {
			return True;
		} else {
			return $response;
		}
	}

	function getEvents($mailbox, $filter = "") {
		if ($filter != "") $filter = str_replace(":", "%3A", $filter);
		if ($filter != "") $filter = str_replace("/", "%2F", $filter);
		if ($filter != "") $filter = "?".str_replace(" ", "%20", $filter);
		$eventList = json_decode($this->sendGetRequest($this->baseURL . 'users/' . $mailbox . '/events'.$filter));
		return($eventList);
	}

    function basicAddress($addresses) {
        $ret = [];
        foreach ($addresses as $address) {
            $ret[] = $address->emailAddress->address;
        }
        return $ret;
    }

    function sendDeleteRequest($URL) {
        echo $URL.PHP_EOL;
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'DELETE');
        curl_setopt($ch, CURLOPT_HTTPHEADER, array('Authorization: Bearer ' . $this->Token, 'Content-Type: application/json'));
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        curl_close($ch);
        return $response;
    }

    function sendPostRequest($URL, $Fields, $Headers = false) {
        echo $URL.PHP_EOL;
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
        echo $URL.PHP_EOL;
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_HTTPHEADER, array('Authorization: Bearer ' . $this->Token, 'Content-Type: application/json'));
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        curl_close($ch);
        return $response;
    }

    function sendPatchRequest($URL, $Fields, $Headers = false) {
        echo $URL.PHP_EOL;
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'PATCH');
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
}
?>
