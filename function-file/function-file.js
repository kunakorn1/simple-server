/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

Office.initialize = function() {
    // TODO: Add your initialization logic here
}
 
function onCommandClick(event) {
    // TODO: Add your command logic here
 
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.
        setAsync("Hello Outlook!", function(asyncResult) {
 
            // Tell the host that we are done
            event.completed();
        });
}
