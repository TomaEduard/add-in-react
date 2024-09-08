import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import {
  Link,
  makeStyles,
  Toast,
  ToastBody,
  Toaster,
  ToastFooter,
  ToastTitle,
  useId,
  useToastController
} from "@fluentui/react-components";
import { DesignIdeas24Regular, LockOpen24Regular, Ribbon24Regular } from "@fluentui/react-icons";
import { createNewSheet, displayJson, insertText, makeCellYellow } from "../taskpane";
import { ToastIntent } from "@fluentui/react-toast";
import Login from "./Login";

export interface AppProps {
  title: string;
}

export interface NotifyFunction {
  (
    status: ToastIntent,
    title: string,
    body: string,
    timeout?: number
  ): void;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    height: "100vh",
    overflow: "hidden",
    boxSizing: "border-box",
    "& *": {
      boxSizing: "border-box"
    }
  }
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration"
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality"
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro"
    }
  ];
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const notify: NotifyFunction = (
    status: ToastIntent = "success",
    title: string = "This is a toast title",
    body: string = "This is a toast body",
    timeout?: number
  ) =>
    dispatchToast(
      <Toast> <ToastTitle action={<Link>Undo</Link>}>{title}</ToastTitle>
        <ToastBody >{body}</ToastBody>
      </Toast>,
      { intent: status, timeout: timeout }
    );

  return (
    <div className={styles.root} style={{ border: "5px solid red" }}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      <Login notify={notify} displayJson={displayJson} />
      <TextInsertion insertText={insertText} makeCellYellow={makeCellYellow} createNewSheet={createNewSheet} notify={notify} />
      <Toaster toasterId={toasterId} />
    </div>
  );
};

export default App;