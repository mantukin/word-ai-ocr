import * as React from "react";
import { makeStyles, Button, Textarea, Field, Text, Title3 } from "@fluentui/react-components";
import { Copy24Regular, Play24Regular, ArrowDownload24Regular, Delete24Regular } from "@fluentui/react-icons";
import { SYSTEM_PROMPT } from "../prompt";
import { generateDocument } from "../document-generator";
import { DocumentStructure } from "../types";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    padding: "15px",
    display: "flex",
    flexDirection: "column",
    gap: "15px",
    minHeight: "100vh",
    backgroundColor: "white",
    boxSizing: "border-box",
  },
  textarea: {
    minHeight: "300px",
    fontFamily: "monospace",
  },
  buttonGroup: {
    display: "flex",
    gap: "10px",
    flexWrap: "wrap"
  }
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [jsonInput, setJsonInput] = React.useState("");
  const [status, setStatus] = React.useState("");

  const copyPrompt = () => {
    navigator.clipboard.writeText(SYSTEM_PROMPT).then(() => {
        setStatus("Prompt copied! Paste it into your AI chat.");
    }).catch(err => {
        setStatus("Failed to copy: " + err);
    });
  };

  const downloadPrompt = () => {
    const element = document.createElement("a");
    const file = new Blob([SYSTEM_PROMPT], {type: 'text/plain'});
    element.href = URL.createObjectURL(file);
    element.download = "ai-instruction.txt";
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
    setStatus("Instruction file downloaded.");
  };

  const handleGenerate = async () => {
    if (!jsonInput) return;
    
    try {
      setStatus("Parsing JSON...");
      // Basic cleanup of JSON (sometimes AI adds markdown blocks)
      let cleanJson = jsonInput.trim();
      if (cleanJson.startsWith("```json")) cleanJson = cleanJson.replace(/^```json/, "").replace(/```$/, "");
      if (cleanJson.startsWith("```")) cleanJson = cleanJson.replace(/^```/, "").replace(/```$/, "");
      
      const structure: DocumentStructure = JSON.parse(cleanJson);
      
      setStatus("Generating document...");
      await generateDocument(structure);
      
      setStatus("Done! Document created.");
    } catch (e) {
      console.error(e);
      setStatus("Error: " + (e as Error).message);
    }
  };

  return (
    <div className={styles.root}>
      <Field label="1. Get Instruction for AI" hint="Send this to ChatGPT/Claude with your image">
        <div className={styles.buttonGroup}>
          <Button 
            icon={<Copy24Regular />} 
            onClick={copyPrompt}
          >
            Copy
          </Button>
          <Button 
            icon={<ArrowDownload24Regular />} 
            onClick={downloadPrompt}
          >
            Download .txt
          </Button>
        </div>
      </Field>

      <Field label="2. Paste AI Response (JSON)">
        <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
          <Textarea 
            className={styles.textarea}
            value={jsonInput}
            onChange={(_e, data) => setJsonInput(data.value)}
            placeholder='Paste the JSON response here...'
            resize="vertical"
          />
          <Button 
            size="small" 
            icon={<Delete24Regular />} 
            onClick={() => { setJsonInput(""); setStatus("Input cleared."); }}
            style={{ alignSelf: 'flex-end' }}
          >
            Clear Input
          </Button>
        </div>
      </Field>

      <Button 
        appearance="primary" 
        size="large"
        icon={<Play24Regular />}
        onClick={handleGenerate}
        disabled={!jsonInput}
      >
        Generate Document
      </Button>

      <Text>{status}</Text>
    </div>
  );
};

export default App;
