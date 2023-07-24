// Contains logic to handle context buttons
/* global document, Office */
const emailContextButton = document.getElementById("emailContextButton");
const senderContextButton = document.getElementById("senderContextButton");
const conversationContextButton = document.getElementById("conversationContextButton");
let activeContext = "mail";

emailContextButton.addEventListener("click", () => {
  // Set the active button
  emailContextButton.classList.add("active");
  senderContextButton.classList.remove("active");
  conversationContextButton.classList.remove("active");

  // loadEmailNote();
});

senderContextButton.addEventListener("click", () => {
  // Set the active button
  emailContextButton.classList.remove("active");
  senderContextButton.classList.add("active");
  conversationContextButton.classList.remove("active");

  // loadSenderNote();
});

conversationContextButton.addEventListener("click", () => {
  // Set the active button
  emailContextButton.classList.remove("active");
  senderContextButton.classList.remove("active");
  conversationContextButton.classList.add("active");

  // loadConversationNote();
});

export function initContextButtons() {
  // Set the initial active button
  emailContextButton.classList.add("active");
}

export function setActiveContext(context: string) {
  switch (context) {
    case "mail":
      emailContextButton.classList.add("active");
      senderContextButton.classList.remove("active");
      conversationContextButton.classList.remove("active");
      activeContext = "mail";
      break;
    case "sender":
      emailContextButton.classList.remove("active");
      senderContextButton.classList.add("active");
      conversationContextButton.classList.remove("active");
      activeContext = "sender";
      break;
    case "conversation":
      emailContextButton.classList.remove("active");
      senderContextButton.classList.remove("active");
      conversationContextButton.classList.add("active");
      activeContext = "conversation";
      break;
    default:
      throw new Error("Invalid context");
  }
}

export function getActiveContext(): string {
  return activeContext;
}

// function loadEmailNote(mailId) {}

// function loadSenderNote(senderId) {}

// function loadConversationNote(conversationId) {}
