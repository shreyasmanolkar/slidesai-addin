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
    //     setAuthUser(responseData);
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

    const initiateGoogleAuth = () => {
      Office.context.ui.displayDialogAsync('https://localhost:5000/google/callback', { height: 30, width: 20 }, (result) => {
        console.log('res', result);
        let dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
          // let authCode = args;
          console.log('args', args);   
        })
      })
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
