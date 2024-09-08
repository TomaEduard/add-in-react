import * as React from "react";
import { useState } from "react";
import { Button, Field, makeStyles, Textarea, tokens } from "@fluentui/react-components";

interface TextInsertionProps {
  insertText: (text: string) => void;
  makeCellYellow: (text: string) => void;
  tooltip: () => void;
  notify: () => void;
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
    props.notify();
  };

  const handleTakeCellYellow = async () => {
    await props.makeCellYellow(text);
    props.notify();
  };
  const handleTooltip = async () => {
    await props.tooltip();
    props.notify();
  };

  // eslint-disable-next-line no-undef
  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return null;
/*
  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Enter text to be inserted into the document.">
        <Textarea size="large" value={text} onChange={handleTextChange} /> </Field>
      <Field className={styles.instructions}>Click the button to insert text.</Field>
      <div className={styles.buttonContainer}>
        <Button className={styles.button} appearance="outline" disabled={false} size="medium" onClick={handleTextInsertion}> Insert text </Button>
        <Button className={styles.button} appearance="outline" disabled={false} size="small" onClick={handleTakeCellYellow}> Make Yellow </Button>
      </div>
    </div>
  );*/

};

export default TextInsertion;