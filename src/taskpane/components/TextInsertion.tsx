import * as React from "react";
import { useState } from "react";
import { Button, Field, makeStyles, Textarea, tokens } from "@fluentui/react-components";
import { NotifyFunction } from "./App";
import { createNewSheet } from "../taskpane";

interface TextInsertionProps {
  insertText: (text: string) => void;
  makeCellYellow: (text: string) => void;
  createNewSheet: (sheetName: string) => void;
  notify: NotifyFunction;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px"
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center"
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%"
  },
  buttonContainer: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    gap: "10px",
    marginTop: "20px"
  },
  button: {
    flex: "1"
  }
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
    await props.insertText(text);
    props.notify("success", "Text Inserted", "The text has been inserted successfully.");
  };

  const handleMakeCellYellow = async () => {
    await props.makeCellYellow(text);
    props.notify("success", "Cell Colored", "The cell has been colored yellow.");
  };

  const handleCreateNewSheet = async () => {
    const sheetName = "NewSheet"; // You can change this to get the sheet name dynamically
    await props.createNewSheet(sheetName);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Enter text to be inserted into the document.">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <div className={styles.buttonContainer}>
        <Button className={styles.button} appearance="outline" size="small" onClick={handleTextInsertion}>
          Insert text
        </Button>
        <Button className={styles.button} appearance="outline" size="small" onClick={handleMakeCellYellow}>
          Make Yellow
        </Button>
        <Button className={styles.button} appearance="outline" size="small" onClick={handleCreateNewSheet}>
          New Sheet
        </Button>
      </div>
      <div className={styles.buttonContainer}>
        <Button className={styles.button} appearance="outline" size="small" onClick={handleTextInsertion}>
          Insert text
        </Button>
        <Button className={styles.button} appearance="outline" size="small" onClick={handleMakeCellYellow}>
          Make Yellow
        </Button>
        <Button className={styles.button} appearance="outline" size="small" onClick={handleCreateNewSheet}>
          New Sheet
        </Button>
      </div>
    </div>
  );
};

export default TextInsertion;