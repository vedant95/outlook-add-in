/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

let list;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    list = document.getElementById("selected-items");

    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      console.log("Event handler added.");
    });
  }
});

export async function run() {
  // Clear the list of previously selected messages, if any.
  clearList(list);

  // Get the subject line and sender's email address of each selected message and log it to a list in the task pane.
  Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      return;
    }

    const selectedItems = asyncResult.value;
    getItemInfo(selectedItems);
  });
}

// Gets the subject line and sender's email address of each selected message.
async function getItemInfo(selectedItems) {
  for (const item of selectedItems) {
    addToList(item.subject);
    // The loadItemByIdAsync method is currently only available to preview in classic Outlook on Windows.
    if (Office.context.diagnostics.platform === Office.PlatformType.PC) {
      await getSenderEmailAddress(item);
    }
  }
}

// Gets the sender's email address of each selected message.
async function getSenderEmailAddress(item) {
  const itemId = item.itemId;
  await new Promise((resolve) => {
    Office.context.mailbox.loadItemByIdAsync(itemId, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(result.error.message);
        return;
      }

      const loadedItem = result.value;
      const sender = loadedItem.from.emailAddress;
      appendToListItem(sender);

      // Unload the current message before processing another selected message.
      loadedItem.unloadAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        resolve();
      });
    });
  });
}

// Clears the list in the task pane.
function clearList(list) {
  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }
}

// Adds an item to a list in the task pane.
function addToList(item) {
  const listItem = document.createElement("li");
  listItem.textContent = item;
  list.appendChild(listItem);
}

// Appends data to the last item of the list in the task pane.
function appendToListItem(data) {
  const listItem = list.lastChild;
  listItem.textContent += ` (${data})`;
}