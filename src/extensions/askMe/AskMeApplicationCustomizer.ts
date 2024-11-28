import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'AskMeApplicationCustomizerStrings';
import { createDirectLine, renderWebChat } from 'botframework-webchat';
import {DirectLine} from 'botframework-directlinejs';

const LOG_SOURCE: string = 'AskMeApplicationCustomizer';

export interface IAskMeApplicationCustomizerProperties {
  testMessage: string;
}

export default class AskMeApplicationCustomizer
  extends BaseApplicationCustomizer<IAskMeApplicationCustomizerProperties> {

    private styleOptions = {
      // Hide upload button.
      hideUploadButton: true,
      accent: "orange",
      botAvatarBackgroundColor: "black",
      botAvatarImage:
        "link",
      userAvatarImage: "https://avatars.githubusercontent.com/u/0000",
      bubbleBackground: "rgba(0, 0, 255, .1)",
      bubbleFromUserBackground: "rgba(0, 255, 0, .1)",
      rootHeight: "80%",
      rootWidth: "20%",
      backgroundColor: "black",
      bubbleBorderColor: "orange",
      bubbleFromUserBorderColor: "red",
      bubbleTextColor: "white",
      bubbleFromUserTextColor: "white",
      sendBoxButtonColor: "orange"
    };
  
    public onInit(): Promise<void> {
      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
  
      let message: string = this.properties.testMessage;
      if (!message) {
        message = '(No properties were provided.)';
      }
      this.addBotImage();
      return Promise.resolve();
    }
  
  
    private addBotImage(): void {
      const chatImage = document.createElement('img');
      chatImage.src = "link";
      chatImage.alt = 'Chat with an agent';
  
      chatImage.style.position = 'fixed';
      chatImage.id = "Ask-Me-Chat";
      chatImage.style.bottom = '20px';
      chatImage.style.right = '20px';
      chatImage.style.width = '50px'; // Adjust the size as needed
      chatImage.style.height = '50px'; // Adjust the size as needed
      chatImage.style.cursor = 'pointer';
      chatImage.style.zIndex = '1000';
      chatImage.style.borderRadius = '50%';
  
      // Add click event listener to the img
      chatImage.addEventListener('click', () => {
        // Define the specific action to be performed on image click
        this.openChatWindow();
      });
  
      document.body.appendChild(chatImage);
    }
  
    private openChatWindow(): void {
      const styleOptions = {
        // Hide upload button.
        hideUploadButton: true,
        accent: "orange",
        botAvatarBackgroundColor: "black",
        botAvatarImage:
          "image link",
        userAvatarImage: "https://avatars.githubusercontent.com/u/0000",
        bubbleBackground: "rgba(0, 0, 255, .1)",
        bubbleFromUserBackground: "rgba(0, 255, 0, .1)",
        rootHeight: "80%",
        rootWidth: "20%",
        backgroundColor: "black",
        bubbleBorderColor: "orange",
        bubbleFromUserBorderColor: "red",
        bubbleTextColor: "white",
        bubbleFromUserTextColor: "white",
        sendBoxButtonColor: "orange"
      };
  
      // Create the chat window container
      const chatWindow = document.createElement('div');
      chatWindow.id = 'chat-window';
      chatWindow.style.position = 'fixed';
      chatWindow.style.bottom = '80px';
      chatWindow.style.right = '20px';
      chatWindow.style.width = styleOptions.rootWidth;
      chatWindow.style.height = styleOptions.rootHeight;
      chatWindow.style.backgroundColor = styleOptions.backgroundColor;
      chatWindow.style.border = '1px solid #ccc';
      chatWindow.style.boxShadow = '0 0 10px rgba(0, 0, 0, 0.1)';
      chatWindow.style.zIndex = '1001';
      chatWindow.style.padding = '10px';
      chatWindow.style.overflow = 'auto';
  
      // Add the chat content
      chatWindow.innerHTML = `
      <div style="width:50%" id="banner"> 
           <h1 style="color:white" >Ask Me Bot</h1>
         </div>
        <div id="webchat" role="main"></div>
      `;
  
      // Add a close button
      const closeButton = document.createElement('button');
      closeButton.innerText = 'Close';
      closeButton.style.position = 'absolute';
      closeButton.style.top = '10px';
      closeButton.style.right = '10px';
      closeButton.addEventListener('click', () => {
        document.body.removeChild(chatWindow);
      });
  
      chatWindow.appendChild(closeButton);
  
      // Append the chat window to the body
      document.body.appendChild(chatWindow);
  
      this.initializeWebChat();
    }
  
    private async initializeWebChat(): Promise<void> {
  
      const tokenEndpointURL = new URL(
        "https://default654463...."
      );
  
      const locale = document.documentElement.lang || "en";
  
      const apiVersion = tokenEndpointURL.searchParams.get("api-version");
      await this.buildChat(locale, apiVersion, tokenEndpointURL);
    }
  
    private async buildChat(locale: string, apiVersion: string | null, tokenEndpointURL: URL): Promise<void> {
      const [directLineURL, token] = await Promise.all([
        fetch(
          new URL(
            `/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`,
            tokenEndpointURL
          )
        )
          .then((response) => {
            if (!response.ok) {
              throw new Error("Failed to retrieve regional channel settings.");
            }
  
            return response.json();
          })
          .then(({ channelUrlsById: { directline } }) => directline),
        fetch(tokenEndpointURL)
          .then((response) => {
            if (!response.ok) {
              throw new Error("Failed to retrieve Direct Line token.");
            }
  
            return response.json();
          })
          .then(({ token }) => token)
      ]);
  
      console.log('directLineURL:', directLineURL);
      console.log('token:', token);
      
      const directLine = new DirectLine({token, domain: new URL("v3/directline", directLineURL).toString()});
  
   
     
  
      const subscription = directLine.connectionStatus$.subscribe({
        next(value) {
          if (value === 2) {
            directLine
              .postActivity({
                value: {
                  localTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                  locale
                },
                name: "startConversation",
                type: "event", 
                from: { id: "user", name: "User" }
              })
              .subscribe();
  
            // Only send the event once, unsubscribe after the event is sent.
            subscription.unsubscribe();
          }
        }
      });
  
      renderWebChat(
        { directLine, locale, styleOptions: this.styleOptions },
        document.getElementById("webchat")
      );
    }
  }