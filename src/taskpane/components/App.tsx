import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { Button, makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular, Button16Filled } from "@fluentui/react-icons";
import insertText from "../excel-office-document";
import { base64Image } from "./base64Image";
import imageToBase64 from 'image-to-base64';
import { sample1 } from "../../template/samplePPTs";
import { aiResponse } from "../../template/sampleAIResponse";
import { secretKey } from "./constants";
import CryptoJS from 'crypto-js';
import OfficeHelpers, { Authenticator } from "@microsoft/office-js-helpers";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = () => {
    const styles = useStyles();
    // const [loggedIn, setLoggedIn] = React.useState(false);
    // const [authUser, setAuthUser] = React.useState(false);

    // const fetchAuthUser = async () => {
    //   try {
    //     const response = await fetch("http://localhost:5000/api/auth/user", {
    //       method: "GET",
    //       credentials: "include",
    //       headers: {
    //         "Content-Type": "application/json",
    //       },
    //     });
    
    //     if (!response.ok) {
    //       console.log("Not properly authenticated");
    //       setLoggedIn(false);
    //       setAuthUser(null);
    //       return;
    //     }
    
    //     const responseData = await response.json();
    
    //     console.log("User: ", responseData);
    //     setLoggedIn(true);
    //      setAuthUser(responseData);
    //   } catch (error) {
    //     console.error("Error fetching auth user:", error);
    //   }
    // };

    // const redirectToGoogleSSO = async () => {
    //   // let timer: NodeJS.Timeout | number | null = null;
    //   const googleLoginURL = "http://localhost:5000/auth/google";
    //   window.open(
    //     googleLoginURL,
    //     "_blank",
    //     "popup=1,width=500,height=600"
    //   );
    
    //   console.log('nw', newWindow);

    //   if (newWindow) {
    //     timer = setInterval(() => {
    //       if (newWindow.closed) {
    //         console.log("Yay we're authenticated");
    //         fetchAuthUser();
    //         if (timer) clearInterval(timer);
    //       }
    //     }, 500);
    //   }
    // };

    const fetchUserDataFromBackend = async (authCode: string) => {
      try {
        const response = await fetch('https://localhost:5000/userdata', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ authCode }),
        });
        if (!response.ok) {
          throw new Error('Failed to fetch user data');
        }
        const userData = await response.json();
          return userData;
      } catch (error) {
        console.error('Error fetching user data:', error);
        throw error;
      }
    }

    const initiateGoogleAuth = () => {
      let dialog;
      try {
        Office.context.ui.displayDialogAsync('https://localhost:5000/auth/google', { height: 65, width: 30 }, (result) => {
          dialog = result.value;
          console.log('res1', result);
          // console.log('res', JSON.stringify(result, null, 2));
          
          // dialog.close();

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async args => {
            // let authCode = args;
            console.log('args', args);
            // console.log('ac', authCode);
            // dialog.close();
            // let authCode = args.message;
            // const userData = await fetchUserDataFromBackend(authCode);
            // console.log('userData', userData);
          })

          // Office.context.ui.messageParent('sample text');

          // dialog.addEventHandler(Office.EventType.DialogEventReceived, async args => {
          //   console.log('args', args);
          // })


        })
      } catch (error) {
        console.log('err', error); 
      }
    }
    
    const initiateMicrosoftAuth = () => {}

    return (
    <div className={styles.root}>
      <h1>SlidesAI</h1>
      <h3>Create beautiful presentation faster.</h3>
      <p>Lorem ipsum, dolor sit amet consectetur adipisicing elit. Voluptate consectetur doloremque dolore repellendus, eveniet soluta laboriosam praesentium enim molestiae iure perferendis adipisci commodi facere? Corporis harum architecto accusamus similique ab!</p>
      {/* <button onClick={handleGoogleLogin} >Google Sign-In</button> */}
      {/* <button onClick={redirectToGoogleSSO}>Google Sign-In</button> */}
      <button onClick={initiateGoogleAuth}>Google Sign-In</button>
      {/* <button onClick={initiateMicrosoftAuth}>Microsoft Sign-In</button> */}
    </div>
  );
}

export default App;
