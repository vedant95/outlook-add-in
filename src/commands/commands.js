/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.log("Office is ready!");
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message.
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete.
//   event.completed();
// }

/**
 * Function to print the IDs of selected emails
 */
// function getMultipleMailIds(event) {
//   try {
//     // Get the current item in the mailbox (email)

//     console.log("Emails Selected", event);

//     const item = Office.context.mailbox.item;

//     if (item) {
//       const id = item.itemId; // Get the unique ID of the email
//       console.log("Selected Email ID:", id);

//       // Display the email ID in a dialog box
//       Office.context.ui.displayDialogAsync(
//         `https://localhost:3000/taskpane.html?id=${id}`,
//         { height: 50, width: 50 }
//       );
//     } else {
//       console.log("No email selected.");
//       Office.context.ui.displayDialogAsync(
//         `https://localhost:3000/taskpane.html?error=NoEmailSelected`,
//         { height: 30, width: 30 }
//       );
//     }
//   } catch (error) {
//     console.error("Error fetching mail ID:", error);
//     Office.context.ui.displayDialogAsync(
//       `https://localhost:3000/taskpane.html?error=${encodeURIComponent(
//         error.message
//       )}`,
//       { height: 30, width: 30 }
//     );
//   }

//   // Indicate the function is complete
//   event.completed();
// }

// Function to get the IDs of all selected emails
function getMultipleMailIds(event) {
  // Use the Office.js API to get selected item IDs
  Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const selectedItems = asyncResult.value; // Array of selected items

      if (selectedItems && selectedItems.length > 0) {
        console.log("Selected Email IDs:");
        selectedItems.forEach((item, index) => {
          console.log(`Email ${index + 1}: ${item.id}`); // Log each email's ID
        });

        // Optional: Display the IDs to the user
        alert(`Selected Email IDs: \n${selectedItems.map((item) => item.id).join("\n")}`);
      } else {
        console.log("No emails selected.");
        alert("No emails selected.");
      }
    } else {
      console.error(`Failed to get selected items: ${asyncResult.error.message}`);
      alert(`Error: ${asyncResult.error.message}`);
    }
  });

  // Indicate that the function has finished execution
  event.completed();
}


// Register the function with Office.
//Office.actions.associate("action", action);
Office.actions.associate("getMultipleMailIds", getMultipleMailIds);
