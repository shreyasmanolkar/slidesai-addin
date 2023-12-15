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

interface AddInProps {
  title: string;
  userInfo: any
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const AddIn = (props: AddInProps) => {
  const styles = useStyles();
  const [heading, setHeading] = React.useState('Default Heading');
  const [description, setDescription] = React.useState('Default Description');
  const [base64Slide, setBase64Slide] = React.useState('');
  
  const [chosenMaster, setChosenMaster] = React.useState('');
  const [chosenLayout, setChosenLayout] = React.useState('');

  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];


  interface SampleSlideProps {
    heading: string;
    description: string;
    onHeadingChange: (value: string) => void;
    onDescriptionChange: (value: string) => void;
  }

  async function fetchData(url) {
    try {
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
  
      const data = await response.json();
      // Handle the JSON data here
      // console.log('JSON Data:', data);
  
      return data;
    } catch (error) {
      // Handle errors here
      console.error('Error fetching data:', error);
      throw error; // Rethrow the error if needed
    }
  }
  
  const SampleSlide: React.FunctionComponent<SampleSlideProps> = (props) => {
    return (
      <div>
        <h2>{props.heading}</h2>
        <label>
          Heading:
          <input type="text" value={props.heading} onChange={(e) => props.onHeadingChange(e.target.value)} />
        </label>
        <br />
        <label>
          Description:
          <textarea value={props.description} onChange={(e) => props.onDescriptionChange(e.target.value)} />
        </label>
        {/* <p>{props.description}</p> */}
      </div>
    );
  };

  const createPowerPointSlide = async () => {
    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;

        // console.log('slides', slides);

        // // const firstSlide = slides.items[0];
        // const firstSlide = slides.getItemAt(0);
        // firstSlide.load("layout, slideMaster");
        
        // await context.sync();
      
        // Retrieve the layout ID and slide master ID
        // const layoutId = firstSlide.layout.id;
        // const slideMasterId = firstSlide.slideMaster.id;
      
        // console.log("Layout ID:", layoutId);
        // console.log("Slide Master ID:", slideMasterId);

        // slides.add({layoutId, slideMasterId});
        slides.add();
        
        // newSlide.title.text = heading;
        // newSlide.content.text = description;

        await context.sync();

        // const shapes = context.presentation.slides.getItemAt(1).shapes;
        // const textbox = shapes.addTextBox("Sample Title!");
        
        // await context.sync();
        
      });

    } catch (error) {
      console.error(error);
    }
  };

  const insertImage = () => {
    Office.context.document.setSelectedDataAsync(
      base64Image, 
      { 
        coercionType: Office.CoercionType.Image,
        // imageLeft: 50,
        // imageTop: 50,
        // imageWidth: 400
      }, 
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log('error', asyncResult.error.message);
        }
      }
    )
  }

  const insertText = () => {
    Office.context.document.setSelectedDataAsync("Hello World!", (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log('error', asyncResult.error.message);
      }
    });
  }

  const getSlideMetadata = () => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log('error', asyncResult.error.message);
      } else {
        console.log("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
      }
    });
  }
  
  async function addSlides() {
    await PowerPoint.run(async function (context) {
      // context.presentation.slides.add();
      
      context.presentation.slides.add();

      // const items = context.presentation.slides.getCount();
      const items = context.presentation.slides.getItemAt(0);

      console.log('i', items);

      await context.sync();

      goToLastSlide();
      console.log("Success: Slides added.");
    });
  }

  async function deleteSlide() {
    // console.log('o', JSON.stringify(Office, null, 2));
    // console.log('o.c.d.s', Office.context.document.settings);

    // console.log('pp', JSON.stringify(PowerPoint, null, 2));

    await PowerPoint.run(async function (context) {
      // console.log('c', context);
      // console.log('c2', JSON.stringify(context, null, 2));

      console.log('c.p.l',  context.presentation.load());
      console.log('c.p.s',  context.presentation.slides);
      console.log('c.p.s',  context.presentation.slideMasters);
      console.log('c.p.tj',  context.presentation.toJSON());

    })
  }
  
  function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Error: " + asyncResult.error.message);
      }
    });
  }
  
  function goToLastSlide() {
    Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Error: " + asyncResult.error.message);
      }
    });
  }
  
  function goToPreviousSlide() {
    Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Error: " + asyncResult.error.message);
      }
    });
  }
  
  function goToNextSlide() {
    Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Error: " + asyncResult.error.message);
      }
    });
  }

  const blankPresentation = () => {
    PowerPoint.createPresentation();
  }

  const createPresentationFromExisting = (event) => {
    try {
      const myFile = event.target.files[0];
      const reader = new FileReader();

      reader.onload = () => {
        const startIndex = reader.result.toString().indexOf("base64,");
        const copyBase64 = reader.result.toString().substr(startIndex + 7);

        // PowerPoint.createPresentation(copyBase64);     
        console.log('cb64', copyBase64);

        setBase64Slide(copyBase64);

      }

      reader.readAsDataURL(myFile);
    } catch (error) {
      console.log('err', error);  
    }
  }

  const insertAllSlides = async () => {
    await PowerPoint.run(async function(context) {
      context.presentation.insertSlidesFromBase64(base64Slide);
      await context.sync();
    });
  }

  const logSlideMasters = async () => {
    await PowerPoint.run(async function(context) {
      const slideMasters = context.presentation.slideMasters.load("id, name, layouts/items/name, layouts/items/id");
  
      await context.sync();
  
      for (let i = 0; i < slideMasters.items.length; i++) {
        console.log("Master name: " + slideMasters.items[i].name);
        console.log("Master ID: " + slideMasters.items[i].id);
        const layoutsInMaster = slideMasters.items[i].layouts;
        for (let j = 0; j < layoutsInMaster.items.length; j++) {
          console.log("    Layout name: " + layoutsInMaster.items[j].name + " Layout ID: " + layoutsInMaster.items[j].id);
        }
      }
    });
  }

  async function addCustomSlide() {
  
    await PowerPoint.run(async function(context) {
      context.presentation.slides.add({
        slideMasterId: chosenMaster,
        layoutId: chosenLayout
      });
      await context.sync();
    });
  }

  async function addImageSlides() {
  
    await PowerPoint.run(async function(context) {
      context.presentation.slides.add({
        slideMasterId: '2147483648#57556370',
        layoutId: '2147483657#1262196759'
      });
      await context.sync();
    });
  }

  async function setSelectedShapes() {
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slide1 = context.presentation.slides.getItemAt(0);
      slide1.load("shapes");
      await context.sync();
      const shapes = slide1.shapes;
      const shape1 = shapes.getItemAt(0);
      const shape2 = shapes.getItemAt(1);
      shape1.load("id");
      shape2.load("id");
      await context.sync();
      slide1.setSelectedShapes([shape1.id, shape2.id]);
      await context.sync();
    });
  }

  async function setSelectedShapeFirst() {
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slide1 = context.presentation.slides.getItemAt(0);
      slide1.load("shapes");
      await context.sync();
      const shapes = slide1.shapes;
      const shape1 = shapes.getItemAt(0);
      shape1.load("id");

      await context.sync();

      console.log('s1', shape1.load("width,height,type,name"));

      await context.sync();
      slide1.setSelectedShapes([shape1.id]);
      
      await context.sync();
    });
  }

  async function setSelectedShapeSecond() {
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slide1 = context.presentation.slides.getItemAt(0);
      slide1.load("shapes");
      await context.sync();
      const shapes = slide1.shapes;
      const shape2 = shapes.getItemAt(1);
      shape2.load("id");
      await context.sync();
      console.log('s2', shape2.load("width,height,type,name"));
      await context.sync();
      slide1.setSelectedShapes([shape2.id]);
      
      await context.sync();
    });
  }

  async function setSelectedShapeThird() {
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slide1 = context.presentation.slides.getItemAt(0);
      slide1.load("shapes");
      await context.sync();
      const shapes = slide1.shapes;
      const shape2 = shapes.getItemAt(2);
      shape2.load("id");
      await context.sync();
      console.log('s3', shape2.load("width,height,type,name"));
      await context.sync();
      slide1.setSelectedShapes([shape2.id]);
      
      await context.sync();
    });
  }

  async function setSelectedShapeNone() {
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slide1 = context.presentation.slides.getItemAt(0);
      slide1.load("shapes");
      await context.sync();
      const shapes = slide1.shapes;
      const shape2 = shapes.getItemAt(1);
      shape2.load("id");
      await context.sync();
      slide1.setSelectedShapes([]);
      
      await context.sync();
    });
  }

  const insertTitle = () => {
    Office.context.document.setSelectedDataAsync("Title On Slides", (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log('error', asyncResult.error.message);
      }
    });
  }

  const insertDescription = () => {
    Office.context.document.setSelectedDataAsync("A slide is a single page of a presentation. Collectively, a group of slides may be known as a slide deck.", (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log('error', asyncResult.error.message);
      }
    });
  }
  
  let savedSlideSelection = [];
  let savedShapeSelection = [];
  async function saveShapeSelection() {
    // Saves which shapes are selected so that they can be reselected later.
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slides = context.presentation.getSelectedSlides();
      const slideCount = slides.getCount();
      slides.load("items");
      await context.sync();
      savedSlideSelection = [];
      slides.items.map((slide) => {
        savedSlideSelection.push(slide.id);
      });
      const shapes = context.presentation.getSelectedShapes();
      const shapeCount = shapes.getCount();
      shapes.load("items");
      await context.sync();
      shapes.items.map((shape) => {
        savedShapeSelection.push(shape.id);
      });
    });
  }

  // const importTheme = () => {
  //   // Load the source presentation
  //   PowerPoint.run(async (context) => {
  //     let sourcePresentation = context.presentation.load("path_to_source_presentation");
  //     await context.sync();

  //     // Get the slide master from the source presentation
  //     let sourceSlideMaster = sourcePresentation.slideMasters.getItemAt(0);

  //     // Add the slide master to the current presentation
  //     let currentPresentation = context.presentation;
  //     currentPresentation.slideMasters.addClone(sourceSlideMaster);

  //     await context.sync();
  //   });
  // }

  const createPPTSlide = async () => {
    // await addSlides();
    await setSelectedShapeFirst();
    await insertTitle();
    await setSelectedShapeSecond();
    await insertDescription();
    await setSelectedShapeNone();
  }

  const createPPTImageSlide = async () => {
    await addImageSlides();
    await setSelectedShapeFirst();
    await insertTitle();
    await setSelectedShapeSecond();
    await insertImage();
    await setSelectedShapeThird();
    await insertDescription();
    await setSelectedShapeNone();
  }

  const slideMasters = async () => {
    await PowerPoint.run(async function(context) {
      const slideMasters = context.presentation.slideMasters.load("id, name, layouts/items/name, layouts/items/id");
  
      await context.sync();
  
      for (let i = 0; i < slideMasters.items.length; i++) {
        console.log("Master name: " + slideMasters.items[i].name);
        console.log("Master ID: " + slideMasters.items[i].id);
        const layoutsInMaster = slideMasters.items[i].layouts;
        for (let j = 0; j < layoutsInMaster.items.length; j++) {
          console.log("    Layout name: " + layoutsInMaster.items[j].name + " Layout ID: " + layoutsInMaster.items[j].id);
        }
      }
    });
  }

  // const addImageFromUrl = () => {
  //   let imageUrl = "https://images.unsplash.com/photo-1686172903880-72c472e79022?q=80&w=1888&auto=format&fit=crop&ixlib=rb-4.0.3&ixid=M3wxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8fA%3D%3D";

  //   let imageData = imageToBase64(imageUrl)
  //   .then(
  //     (response) => {
  //       console.log('res', response);
  //       return response;
  //     }
  //   )
  //   .catch(
  //     (error) => {
  //       console.log(error);
  //     }
  //   );

  //   Office.context.document.setSelectedDataAsync(
  //     imageData, 
  //     { 
  //       coercionType: Office.CoercionType.Image,
  //     }, 
  //     (asyncResult) => {
  //       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
  //         console.log('err', asyncResult.error.message);
  //       }
  //     }
  //   )
  // }

  // const onlinePres = async () => {
  //   try {
  //     const response = await fetch("https://1drv.ms/p/s!Av3uzvkUfrhyhVaM0-tgt74cv-2L?e=qlhFrh", { mode: 'no-cors' });

  //     console.log('res', response);

  //     const pres = await response.json();
      
  //     console.log('pres', pres);
  //   } catch (error) {
  //     console.error('Error fetching presentation:', error);
  //   }
  // }

  const createSampleTheme = async () => {
    await PowerPoint.run(async function(context) {
      context.presentation.insertSlidesFromBase64(sample1);
      await context.sync();

      const count = context.presentation.slides.getCount();
      
      await context.sync();

      for(let i = 1; i <= count.value; i++) {
        const slide = context.presentation.slides.getItemAt(i);
        slide.delete();
      }

    });
  }

  const addTitleSlide = async () => {
    await PowerPoint.run(async function(context) {
      context.presentation.slides.add({
        slideMasterId: '2147483684#0',
        layoutId: '2147483685#0'
      });
      await context.sync();
    });
  }

  const addTitleAndContentSlide = async () => {
    await PowerPoint.run(async function(context) {
      context.presentation.slides.add({
        slideMasterId: '2147483684#0',
        layoutId: '2147483686#0'
      });
      await context.sync();
    });
  }

  const addThankyouSlide = async () => {
    await PowerPoint.run(async function(context) {
      context.presentation.slides.add({
        slideMasterId: '2147483684#0',
        layoutId: '2147483690#0'
      });
      await context.sync();
    });
  };

  function decryptData(encryptedData, secretKey) {
    try { 
      const decrypted = CryptoJS.AES.decrypt(encryptedData, secretKey);
      let decryptedData = JSON.parse(decrypted.toString(CryptoJS.enc.Utf8));
      return decryptedData;
    } catch (error) {
      console.error("Error decrypting data:", error.message);
      return null;
    }
  }
  
  const createThemePPT = async () => {
    try {
      const aiResponseData = await fetchData('https://localhost:5000/v1/ai-response');
      const themeDataData = await fetchData('https://localhost:5000/v1/base64');

      const aiResponse = decryptData(aiResponseData.ed, secretKey);
      const themeData = decryptData(themeDataData.ed,  secretKey);

      let count = 0;

      await PowerPoint.run(async function(context) {

        context.presentation.insertSlidesFromBase64(themeData.base64);
        await context.sync();

        const slideCount = context.presentation.slides.getCount();
        await context.sync();

        count = slideCount.value;
      });
      
      await addTitleSlide();
      await setSelectedShape(count, 0);
      await insertThemeTitle(aiResponse.title);
      await setSelectedShape(count, 1);
      await insertThemeDescription(aiResponse.desc);
      await addTitleAndContentSlide();
      await setSelectedShape(count + 1, 0);
      await insertThemeTitle(aiResponse.slidePoints[0].title);
      await setSelectedShape(count + 1, 1);
      await insertThemeDescription(aiResponse.slidePoints[0].descriptions[0]);
      await addTitleAndContentSlide();
      await setSelectedShape(count + 2, 0);
      await insertThemeTitle(aiResponse.slidePoints[1].title);
      await setSelectedShape(count + 2, 1);
      await insertThemeDescription(aiResponse.slidePoints[1].descriptions[0]);
      await addThankyouSlide();
      await setSelectedShape(count + 3, 0);
      await insertThemeTitle(aiResponse.thankYouText);

      await PowerPoint.run(async function(context) {      
        await context.sync();
        for(let i = 0; i < count; i++) {
          const slide = context.presentation.slides.getItemAt(i);
          slide.delete();
        }
      });

      await goToFirstSlide();

    } catch (error) {
      console.log('e', error);
    }
  }

  async function setSelectedShape(slideNumber, shapeNumber) {
    await PowerPoint.run(async (context) => {
      context.presentation.load("slides");
      await context.sync();
      const slide = context.presentation.slides.getItemAt(slideNumber);
      
      // console.log('slide', slide);
      
      slide.load("shapes");
      await context.sync();
      const shapes = slide.shapes;
      
      // console.log('shapes', shapes);

      const shape = shapes.getItemAt(shapeNumber);
      shape.load("id");
      await context.sync();
      // console.log('s2', shape.load("width,height,type,name"));
      await context.sync();
      slide.setSelectedShapes([shape.id]);
      
      await context.sync();
    });
  }

  const insertThemeTitle = (title) => {
    Office.context.document.setSelectedDataAsync(title, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log('error', asyncResult.error.message);
      }
    });
  }

  const insertThemeDescription = (desc) => {
    Office.context.document.setSelectedDataAsync(desc, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log('error', asyncResult.error.message);
      }
    });
  }

  const insertTextWithoutBullet = async () => {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      slide.load('shapes');
      await context.sync();
      const shapes = slide.shapes;
      const selectedShape = shapes.getItemAt(1);
      selectedShape.load('textFrame');
      await context.sync();
      const textFrameSS = selectedShape.textFrame;
      textFrameSS.textRange.load('text');
      await context.sync();
      textFrameSS.textRange.text = "The CSS nesting module defines a syntax for nesting selectors, providing the ability to nest one style rule inside another, with the selector of the child rule relative to the selector of the parent rule.";
      await context.sync();
      const textRange = textFrameSS.textRange;
      textRange.load('paragraphFormat');
      await context.sync();
      const paragraphFormat = textRange.paragraphFormat; 
      paragraphFormat.load('bulletFormat');
      await context.sync();
      const bulletFormat = paragraphFormat.bulletFormat;
      bulletFormat.visible = false;
      await context.sync();
    });
  }

  const insertTextWithoutBulletFormat = async () => {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      slide.load('shapes');
      await context.sync();
      const shapes = slide.shapes;
      const selectedShape = shapes.getItemAt(1);
      selectedShape.load('textFrame');
      await context.sync();
      const textFrameSS = selectedShape.textFrame;
      textFrameSS.textRange.load('text');
      
      console.log('tr', textFrameSS.textRange);
      
      await context.sync();
      textFrameSS.textRange.text = "The CSS nesting module defines a syntax for nesting selectors, providing the ability to nest one style rule inside another, with the selector of the child rule relative to the selector of the parent rule.";
      await context.sync();
      const textRange = textFrameSS.textRange;
      textRange.load('paragraphFormat');
      await context.sync();
      const paragraphFormat = textRange.paragraphFormat; 
      
      console.log('pf', paragraphFormat);

      // paragraphFormat.load('bulletFormat');
      // await context.sync();
      // const bulletFormat = paragraphFormat.bulletFormat;
      // bulletFormat.visible = false;
      // await context.sync();
    });
  }

  const getJSONData = () => {
    fetch('https://localhost:5000/v1/ai-response')
      .then(response => {
        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status}`);
        }
        return response.json();
      })
      .then(data => {
        // Handle the JSON data here
        console.log('JSON Data:', data);
        return data;
      })
      .catch(error => {
        // Handle errors here
        console.error('Error fetching data:', error);
      });
  };
  
  const getBase64Data = () => {
    fetch('https://localhost:5000/v1/base64')
      .then(response => {
        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status}`);
        }
        return response.json();
      })
      .then(data => {
        // Handle the JSON data here
        console.log('JSON Data:', data);
        // return data;
      })
      .catch(error => {
        // Handle errors here
        console.error('Error fetching data:', error);
      });
  };
  
  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      {/* <HeroList message="Discover what this add-in can do for you today!" items={listItems} /> */}
      {/* <TextInsertion /> */}
      {/* <SampleSlide heading={heading}
        description={description}
        onHeadingChange={setHeading}
        onDescriptionChange={setDescription}/> */}
  
      <button onClick={insertImage}>
        insert Image 
      </button>
  
      <button onClick={insertText}>
        insert Text
      </button>
  
      <button onClick={getSlideMetadata}>
        Get Slide Metadata
      </button>
  
      <button onClick={goToNextSlide}>
        Go To Next Slide
      </button>
      
      <button onClick={goToPreviousSlide}>
        Go To Previous Slide
      </button>
  
      <button onClick={goToLastSlide}>
        Go To Last Slide
      </button>
    
      <button onClick={goToFirstSlide}>
        Go To First Slide
      </button>
      
      <button onClick={addSlides}>
        add slides
      </button>
  
      <button onClick={deleteSlide}>
        sample
      </button>
  
      <button onClick={blankPresentation}>
        Blank Presentation
      </button>
  
      <form>    
        <input type="file" id="file" onChange={createPresentationFromExisting} />
      </form>
  
      <button onClick={logSlideMasters}>
        Log Slide Masters
      </button>
         
      <form>   
        <label htmlFor="master">Master</label>
        <input name="master" onChange={e => setChosenMaster(e.target.value)} />
        <br />
        <label htmlFor="layout">Layout</label>
        <input name="layout" onChange={e => setChosenLayout(e.target.value)} />
      </form>
  
      <button onClick={addCustomSlide}>
        add custom slide
      </button>
  
      <button onClick={setSelectedShapes}>
        Select First Two Shapes
      </button>
  
      <button onClick={setSelectedShapeFirst}>
        Select First Shape
      </button>
  
      <button onClick={setSelectedShapeSecond}>
        Select Second Shape
      </button>

      <button onClick={setSelectedShapeThird}>
        Select Third Shape
      </button>
  
      <button onClick={createPPTSlide}>
        Create PPT Slide
      </button>
  
      <button onClick={addImageSlides}>
        add image slides
      </button>
 
      <button onClick={createPPTImageSlide}>
        create Image PPT slides
      </button>

      <button onClick={slideMasters}>
        Slide Masters
      </button>

      <button onClick={insertAllSlides}>
        Insert All Slides
      </button>

      {/* <button onClick={onlinePres}>
        press
      </button> */}

      <button onClick={createSampleTheme}>
        create Slides with Sample Theme
      </button>

      <button onClick={addTitleSlide}>
        Add theme Title Slide
      </button>

      <hr />

      <button onClick={getJSONData}>
        Get AI Response
      </button>

      <button onClick={getBase64Data}>
        Get Theme
      </button>

      <button onClick={createThemePPT}>
        Add Theme Slide
      </button>

      <button onClick={insertTextWithoutBullet}>
        Insert Text Without Bullet
      </button>

      <button onClick={insertTextWithoutBulletFormat }>
        Insert Text Without Bullet Format
      </button>
    </div>
  );
}

export default AddIn;
