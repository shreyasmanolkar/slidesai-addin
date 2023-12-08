import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import AddIn from "./AddIn";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = () => {
    const styles = useStyles();
    const [loggedIn, setLoggedIn] = React.useState(false);
    const [authUser, setAuthUser] = React.useState(null);

    React.useEffect(() => {
      checkAndSetAuthUser();
    }, []); 

    const checkAndSetAuthUser = () => {
      const storedUserInfo = localStorage.getItem('userInfo');
    
      if (storedUserInfo) {
        const userInfo = JSON.parse(storedUserInfo);
        setAuthUser(userInfo);
        setLoggedIn(true);
      }
    };

    const initiateGoogleAuth = () => {
      let dialog;
      try {
        Office.context.ui.displayDialogAsync('https://localhost:5000/auth/google', { height: 65, width: 30 }, (result) => {
          dialog = result.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived , async args => {
            // dialog.close();
            const userInfo = JSON.parse(args.message);
            // console.log('ui', userInfo);

            localStorage.setItem('userInfo', JSON.stringify(userInfo));
            
            setAuthUser(userInfo);
            setLoggedIn(true);
          })
        })
      } catch (error) {
        console.log('err', error); 
      }
    }
    
    const initiateMicrosoftAuth = () => {}

    return ( loggedIn ? <AddIn title={"Contoso Task Pane Add-in"} userInfo={authUser}/> :
    <div className={styles.root}>
      <h1>SlidesAI</h1>
      <h3>Create beautiful presentation faster.</h3>
      <p>Lorem ipsum, dolor sit amet consectetur adipisicing elit. Voluptate consectetur doloremque dolore repellendus, eveniet soluta laboriosam praesentium enim molestiae iure perferendis adipisci commodi facere? Corporis harum architecto accusamus similique ab!</p>
      <button onClick={initiateGoogleAuth}>Google Sign-In</button>
      {/* <button onClick={initiateMicrosoftAuth}>Microsoft Sign-In</button> */}
    </div>
  );
}

export default App;
