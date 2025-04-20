import logo from './logo.svg';
import './App.css';
import Slider from '@mui/material/Slider';
import { styled } from '@mui/material/styles';
// import Button from '@mui/material/Button';
import Tooltip, { tooltipClasses } from '@mui/material/Tooltip';
import InfoIcon from '@mui/icons-material/Info';
import IconButton from '@mui/material/IconButton';
import React, { useState,useEffect } from 'react';
import axios from 'axios';
import { Button, Snackbar } from '@mui/material';
// import Accordion from '@mui/material/Accordion';
// import AccordionSummary from '@mui/material/AccordionSummary';
// import AccordionDetails from '@mui/material/AccordionDetails';
// import Typography from '@mui/material/Typography';
// import ArrowDownwardIcon from '@mui/icons-material/ArrowDownward';
// import ArrowDropDownIcon from '@mui/icons-material/ArrowDropDown';


function App() {
  const marks = [
    {
      value: 0,
      label: '1',
    },
    {
      value: 25,
      label: '2',
    },
    {
      value: 50,
      label: '3',
    },
    {
      value: 75,
      label: '4',
    },
    {
      value: 100,
      label: '5',
    }
  ];

  const [snack, setSnack] = useState({
    open: false,
    message: '',
    bgColor: '',
  });

  const showSnackbar = (type,message) => {
    let colorMap = {
      success: '#4caf50',
      warning: '#ff9800',
      error: '#f44336',
      info: '#2196f3',
    };

    setSnack({
      open: true,
      message: message,
      bgColor: colorMap[type],
    });
  };

  const handleClose = () => {
    setSnack((prev) => ({ ...prev, open: false }));
  };

  const longText = `Set creativity level according to your need , lesser level (0-2) maintains the same words and corrects only
  grammer , higher creativity levels improvises more proficiency while keeping same context.`;
  function valueLabelTooltip(value) {
    switch (value) {
      case 0:
        return 'Proof Reader';
      case 25:
        return 'Clarity Refiner';
      case 50:
        return 'Tone Enhancer';
      case 75:
        return 'Message Polisher';
      case 100:
        return 'Rewrite';
      default:
        return '';
    }
  }
  
  const CustomWidthTooltip = styled(({ className, ...props }) => (
    <Tooltip {...props} classes={{ popper: className }} />
  ))({
    [`& .${tooltipClasses.tooltip}`]: {
      maxWidth: 300,
    },
  });


  const [sliderValue, setSliderValue] = useState(0);

  const handleSliderChange = (event, newValue) => {
    setSliderValue(newValue);
    console.log('Current slider value:', newValue);
  };

  const [input, setInput] = useState();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const refineText = async () => {
    setLoading(true);
    setError('');
    if (input === undefined || input === '') {
      setError('"Please write a mail first"');
      // console.log('Empty response from API');
      showSnackbar('warning',"Please write a mail first")
      setLoading(false);
      return;
    }

    try {

      const creativity_level = Number(sliderValue); // Ensure it's a number
      let prompt = '';
      
      switch (creativity_level) {
        case 0:
          prompt = `Proofread the following email. Only fix grammar, punctuation, and spelling errors. Do not change the sentence structure or wording unless absolutely necessary:\n\n"${input}"\n\nReturn the refined email.`;
          break;
        case 25:
          prompt = `Refine the following email by improving its clarity and readability. Do not change the tone or intended meaning. Keep the structure intact and avoid rewording unless it enhances clarity:\n\n"${input}"\n\nReturn the improved version of the email.`;
          break;
        case 50:
          prompt = `Improve the tone and phrasing of the following email while keeping the original meaning and structure intact. Aim for a more professional, polite, and engaging tone:\n\n"${input}"\n\nReturn the refined email.`;
          break;
        case 75:
          prompt = `Polish the following email to enhance its tone, flow, and structure. You may slightly reword or rearrange sentences for better readability, while preserving the original message:\n\n"${input}"\n\nReturn the refined email.`;
          break;
        case 100:
          prompt = `Rewrite the following email completely, keeping the core message and intent the same. Make it sound professional, well-structured, and impactful:\n\n"${input}"\n\nReturn the rewritten email.`;
          break;
        default:
          prompt = `Analyse the sentiment of the following mail : \n\n"${input}"\n\nReturn the sentiment.`;
          break;
      }
      
      // console.log(prompt);
      // await sleep(2000);it
      const response = await axios.post(
        ' https://openrouter.ai/api/v1/chat/completions',
        {
          model: 'gpt-3.5-turbo',
          messages: [{ role: 'user', content: prompt }],
          temperature: 0.7,
        },
        {
          headers: {
            // Authorization: `Bearer ${process.env.REACT_APP_OPENAI_API_KEY}`,
            Authorization: `Bearer ${"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"}`,
            'Content-Type': 'application/json',
          },
          
        }
        
      );
    //   console.log('API KEY:', process.env.REACT_APP_OPENAI_API_KEY);

      const refined = response.data.choices[0].message.content.trim();
      console.log('Refined Text:', refined);
      if (input === undefined || input === '') {
        setError('"Please write a mail first"');
        // console.log('Empty response from API');
        showSnackbar('warning',"Please write a mail first")
        setLoading(false);
        // return;
      }
      else {
        setInput(refined);

        showSnackbar('success',"Successfully refined mail")
       
        
      }
      // console.log(input)
      // Check if the response is empty or not

    //   setInput(refined);
    } catch (err) {
      console.error(err);
      showSnackbar('error',"Something went wrong. Please try again.")
      setError('Something went wrong. Please try again.');
    }

    setLoading(false);
  };



  useEffect(() => {
    // Ensure Office is ready
 
    if (window.Office) {
      // console.log(window)
      window.Office.onReady((info) => {
        if (info.host === window.Office.HostType.Outlook) {
          // Automatically load body on ready
          getMailBody();
        }
      });
    }
  }, []);

  const getMailBody = () => {
    console.log("here")
    window.Office.context.mailbox.item.body.getAsync(
      "text", // or 'html' depending on what you want
      function (asyncResult) {
         
        if (asyncResult.status === window.Office.AsyncResultStatus.Succeeded) {
          const body = asyncResult.value;
          setInput(body); // Set mail body to textarea input
        } else {
          console.error(asyncResult.error.message);
        }
      }
    );
  };

  return (
    <div className="App">

<Snackbar
        open={snack.open}
        onClose={handleClose}
        autoHideDuration={5000}
        message={snack.message}
        anchorOrigin={{ vertical: 'top', horizontal: 'right' }}
        ContentProps={{
          sx: {
            backgroundColor: snack.bgColor,
            color: '#fff',
            fontWeight: 'bold',
          },
        }}
      />
     
     <link href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,100..900;1,14..32,100..900&display=swap" rel="stylesheet"/>
 
{/* <div class="container">
  <div class="grid-features">
    <div class="bento-card cloud">
      <div class="bento-card-description">
        <h2>Refine Email</h2>
        <p>Use a pre-designed template or personalize with video, stickers, fonts, and more</p>
      </div>

      <a class="btn" href="#">
      <span class="text">Refine My Mail</span>
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4.66669 11.3334L11.3334 4.66669" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/><path d="M4.66669 4.66669H11.3334V11.3334" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </a>
    </div> 
    <div class="bento-card logo">
      <div class="bento-card-description">
        <h2>Impact The Environment</h2>
        <p>We built smart solutions for you and the children of tomorrow. All your data will be stored on smart storage</p>
      </div>
      <div class="bento-card-details logo">
      </div>
    </div> 
    <div class="bento-card device">
      <div class="bento-card-description">

      </div>
    </div>
    <div class="bento-card inbox">
      <div class="bento-card-description">
        <h3>Inbox</h3>
        <p>Track your gifts, group chats, and sent cards</p>
      </div>
      <div class="bento-card-details inbox">

      </div>
    </div>
    <div class="bento-card device-2">
    </div>
    <div class="bento-card ai-gen">
      <div class="bento-card-description">
        <h2>AI Generates Your Route</h2>
      </div>
      <div class="bento-card-details ai-gen">

      </div>
    </div>
  </div>
  <div class="credit">
    <a href="https://emilandersson.com/blog/how-to-design-bento-grids">Created for this blog post</a>
  </div>
</div> */}
     <div id="wrapper">
     <div className='col' style={{display :"block",marginTop:"1vh"}}>

     <h2 class="heading" id="padleft"><span style={{color:'#154633'}}>Write </span><span style={{color:"#95C11F"}}>Right!</span></h2>
    <p id="padleft">AI powered mail writing assistant</p>
    </div>
<div class="col">
     
 
<textarea    placeholder="Please draft a mail or Tell me what to write for you" rows="20" name="comment[text]" id="comment_text" cols="40" class={loading ? 'skeleton' : ''} value={input} onChange={(e) => setInput(e.target.value)} autocomplete="off" role="textbox" aria-autocomplete="list" aria-haspopup="true"></textarea>
      
      </div>
     {/* <div class="gallery-template-item">
  <div class="gallery-animated-background">
    <div class="main-masks"></div>
     
  </div>
</div> */}
      
     {/* <div id="textwrapper">
      <textarea name="" id="lined" cols="30" rows="10">
Hello Test
I would like this to be on line.
Very line.</textarea>
      </div> */}
     <div className='col' style={{display: "flex", alignItems: 'center', justifyContent: 'center', flexDirection: 'column'}}>
     {/* <Accordion id= "accordion">
        <AccordionSummary
          
          aria-controls="panel1-content"
          id="panel1-header"
        >
          <Typography component="span">Accordion 1</Typography>
        </AccordionSummary>
        <AccordionDetails>
        <Slider
        aria-label="Restricted values"
        defaultValue={20}
        getAriaValueText={valuetext}
        step={null}
        valueLabelDisplay="auto"
        marks={marks}
        id="slider"
      />
        </AccordionDetails>
      </Accordion> */}
      <div style={{display:"flex" ,alignItems: "center" }} id="creativity"> <h1   >Creativity</h1>  <CustomWidthTooltip title={longText}>
      <IconButton>
      <InfoIcon />
      </IconButton>
</CustomWidthTooltip> </div> 
<Slider
  aria-label="Creativity Level"
  defaultValue={0}
  value={sliderValue}
  onChange={handleSliderChange}
  step={null}
  marks={marks}
  valueLabelDisplay="auto"
  valueLabelFormat={valueLabelTooltip} // Controls tooltip only
  getAriaValueText={(value) => `${value}`} // For accessibility (optional)
  id="slider"
/>

      
      </div>
      
  <div class="col">
    <a   onClick={() => refineText()} class={loading ? 'btn inprogress': 'btn'}>
      <span class="text">Refine My Mail</span>
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4.66669 11.3334L11.3334 4.66669" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/><path d="M4.66669 4.66669H11.3334V11.3334" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </a>
    
  </div>
  {/* <div class="col">
    <a class="btn light" href="#">
      <span class="text">I'm a lovely button</span>
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4.66669 11.3334L11.3334 4.66669" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/><path d="M4.66669 4.66669H11.3334V11.3334" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </a>
  </div> */}
</div>
    </div>
  );
}

export default App;
