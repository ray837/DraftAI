 
import './App.css';
import Slider from '@mui/material/Slider';
import { styled } from '@mui/material/styles';
import Tooltip, { tooltipClasses } from '@mui/material/Tooltip';
import InfoIcon from '@mui/icons-material/Info';
import IconButton from '@mui/material/IconButton';
import React, { useState,useEffect,useRef  } from 'react';
import axios from 'axios';
import { Button, Snackbar } from '@mui/material';
import Fab from '@mui/material/Fab';
 
 
import DriveFileMoveRtlIcon from '@mui/icons-material/DriveFileMoveRtl';

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

  const longText = `Control creativity intensity — low levels apply minimal changes with grammatical corrections, high levels enable advanced rephrasing for polished, professional output.`;
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
      fontSize: '12px',
    },
  });


  const [sliderValue, setSliderValue] = useState(0);

  const handleSliderChange = (event, newValue) => {
    setSliderValue(newValue);
   
  };

  const [input, setInput] = useState();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [toggle, Settoggle] = useState(false);
  const typeText = async (text) => {
    setInput(''); // Clear textarea before typing
    for (let i = 0; i < text.length; i++) {
      await new Promise((resolve) => setTimeout(resolve, 20)); // Adjust delay
      setInput((prev) => prev + text.charAt(i));
    }
  };
  const refineText = async () => {
    setLoading(true);
    setError('');
    if (input === undefined || input === '') {
      setError('"Please write a mail first"');
    
      showSnackbar('warning',"Please write a mail first")
      setLoading(false);
      return;
    }

    try {

      const creativity_level = Number(sliderValue); // Ensure it's a number
      let prompt = '';
      
      switch (creativity_level) {
        case 0:
          prompt = `Proofread the following email. Only fix grammar, punctuation, and spelling errors. Do not change the sentence structure or wording unless absolutely necessary:\n\n"${input}"\n\nReturn the refined email body only.`;
          break;
        case 25:
          prompt = `Refine the following email by improving its clarity and readability. Do not change the tone or intended meaning. Keep the structure intact and avoid rewording unless it enhances clarity:\n\n"${input}"\n\nReturn the improved version of the email body only.`;
          break;
        case 50:
          prompt = `Improve the tone and phrasing of the following email while keeping the original meaning and structure intact. Aim for a more professional, polite, and engaging tone:\n\n"${input}"\n\nReturn the refined email body only.`;
          break;
        case 75:
          prompt = `Polish the following email to enhance its tone, flow, and structure. You may slightly reword or rearrange sentences for better readability, while preserving the original message:\n\n"${input}"\n\nReturn the refined email body only.`;
          break;
        case 100:
          prompt = `Rewrite the following email completely, keeping the core message and intent the same. Make it sound professional, well-structured, and impactful:\n\n"${input}"\n\nReturn the refined email body only.`;
          break;
        default:
          prompt = `Analyse the sentiment of the following mail : \n\n"${input}"\n\nReturn the sentiment.`;
          break;
      }
       
      const response = await axios.post(
      
          "http://127.0.0.1:5000/proxy",
        {
       
          messages: [{ role: 'user', content: prompt }],
           model:'/app/models/gemma-3-27b-it',
    
        },
        {
          headers: {
            
            'Content-Type': 'application/json',
            
          },
          
        }
        
      );
 
 
      const refined = response.data.choices[0].message.content.trim();
      
      console.log('Refined Text:', refined);
      if (input === undefined || input === '') {
        setError('"Please write a mail first"');
      
        showSnackbar('warning',"Please write a mail first")
        setLoading(false);
      
      }
      else {
        setLoading(false);
        await  typeText(refined);
        setInput(refined);

        showSnackbar('success',"Successfully refined mail")
        Settoggle(true)
       
        
      }
       
    } catch (err) {
       
      console.error(err);
      showSnackbar('error',"Something went wrong. Please try again.")
      setError('Something went wrong. Please try again.');
      
    }

    setLoading(false);
  };



  const previousBodyRef = useRef("");

  useEffect(() => {
     
    if (window.Office) {
    
      window.Office.onReady((info) => {
       
        if (info.host === window.Office.HostType.Outlook) {
        
          startPollingBody(); // polling body
        }
      });
    }
  }, []);

  const startPollingBody = () => {
    const interval = setInterval(() => {
      
      window.Office.context.mailbox.item.body.getAsync("text", (result) => {
        if (result.status === window.Office.AsyncResultStatus.Succeeded) {
          const body = result.value;
           
          if (body !== previousBodyRef.current) {
            previousBodyRef.current = body;
            // console.log("inner")
            setInput(body);
          }
        }
      });
    }, 800); 
    return () => clearInterval(interval);
  };


  function insertInputToBody() {
     
    if (input==undefined){
      showSnackbar('error',"No message to insert")
      return
    }else if(input.trim().length==0){
      showSnackbar('error',"No message to insert")
      return
    }

    var subject="";
    var processedinput=""
    if (input.toLowerCase().includes("subject:")) {
     subject = input.split("\n",2)[0];
     console.log(subject)
    
    processedinput=input.split("\n").slice(1).join("\n")
     console.log(processedinput)
    }else{
      processedinput=input
    }
    const htmlFormattedInput = processedinput.replace(/\n/g, "<br/>");
    try{
      window.Office.context.mailbox.item.subject.setAsync(subject);

    window.Office.context.mailbox.item.body.setAsync(
      htmlFormattedInput,
      { coercionType: "html" },
      function (result) {
        if (result.status === window.Office.AsyncResultStatus.Succeeded) {
       
          showSnackbar('success',"Inserted to mail draft")
        } else {
          showSnackbar('error',result.error.message)
        
        }
      }
    
    );}catch{
      showSnackbar('error',"Issue inserting ! This feature only works with Outlook")
      
      navigator.clipboard.writeText(input)
    .then(() => {
      showSnackbar('success', "Feature only works with Outlook.Copied to clipboard instead");
    })
    .catch(() => {
      showSnackbar('error', "Failed to insert and copy to clipboard.");
    });
    }
  }

  return (
    <div className="App">
     
     <link href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,100..900;1,14..32,100..900&display=swap" rel="stylesheet"/>
 
     <div id="wrapper">
     {toggle && (
      <div className='col' style={{margin:"auto"}} onClick={()=>insertInputToBody()}> 
     <Tooltip title="Insert Draft">
    <Fab size="medium"   aria-label="add" style={{
    position: "fixed",
    bottom: "6vw",
    right: "6vw",
    backgroundColor:"rgb(149, 193, 31)",
    color:"#154633",
    zIndex: 1000, 
  }}>
  <DriveFileMoveRtlIcon/>
</Fab>
</Tooltip> 
</div>
)}
     <div className='col' style={{display :"block",marginTop:"2vh"}}>
     <h2 class="heading" id="padleft"><span style={{color:'#154633'}}>Write </span><span style={{color:"#95C11F"}}>Right!</span></h2>
     </div>
    <p id="padleft" style={{fontSize:"2vh"}}>AI-Powered Email Writing Assistant</p>
   
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
            // fontWeight: 'bold',
          },
        }}
       
      />
<div class="col" style={{marginTop:"2vh"}}>
 <textarea placeholder="Draft a mail or let me know what i can draft for you." rows="20" name="comment[text]" id="comment_text" cols="40" class={loading ? 'skeleton' : ''} value={input} onChange={(e) => setInput(e.target.value)} autocomplete="off" role="textbox" aria-autocomplete="list" aria-haspopup="true"></textarea>   
      </div>
     
     <div className='col' style={{display: "flex", alignItems: 'center', justifyContent: 'center', flexDirection: 'column',marginTop:'3vh'}}>
     
      <div style={{display:"flex" ,alignItems: "center" }} id="creativity"> <h1 >Creativity</h1>  <CustomWidthTooltip title={longText} >
      <IconButton sx={{ width: 24, height: 24, paddingX:1.7,paddingY:2 }}>
      <InfoIcon sx={{ fontSize: 16 }}/>
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
  valueLabelFormat={valueLabelTooltip}  
  getAriaValueText={(value) => `${value}`} 
  id="slider"
  sx={{
    height: 4, 
    '& .MuiSlider-thumb': {
      width: 12,
      height: 12,
    },
    '& .MuiSlider-valueLabel': {
      fontSize: '10px',
      // backgroundColor: 'primary.main',
      padding: '2px 6px',
    },
  }}
/>

      
      </div>
      
  <div class="col">
    <a   onClick={() => refineText()} class={loading ? 'btn inprogress': 'btn'}>
      <span class="text">Improve Writing</span>
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4.66669 11.3334L11.3334 4.66669" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/><path d="M4.66669 4.66669H11.3334V11.3334" stroke="white" stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </a>
    
  </div>
  
</div>
    </div>
  );
}

export default App;
