import React from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../../authConfig";
import { AccountInfo } from "@azure/msal-common";
import { AuthenticationResult } from "@azure/msal-browser";
import { NotifyFunction } from "./App";

interface LoginProps {
  notify: NotifyFunction;
  displayJson: (o: object) => void;
}

const Login: React.FC<LoginProps> = ({ notify, displayJson }) => {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance
    .loginPopup(loginRequest)
    .then((result: AuthenticationResult) => {
      const accountInfo: AccountInfo = result.account;
      notify("success", "Login successful", `Hello ${accountInfo.name.split(" ")[0]}!`);
      displayJson(accountInfo);
    })
    .catch((error) => {
      const _error = JSON.stringify(error);
      if (error.errorCode === "user_cancelled") {
        notify("info", "Login failed", `Login failed: ${_error}`);
      } else {
        notify("error", "Login failed", `Login failed: ${_error}`);
      }
    });
  };

  return (
    <div style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
      <img
        src="assets/ms-symbollockup_signin_light.png"
        alt="Login with Microsoft"
        onClick={handleLogin}
        style={{
          cursor: "pointer",
          transition: "opacity 0.3s"
        }}
        onMouseEnter={(e) => (e.currentTarget.style.opacity = "0.74")}
        onMouseLeave={(e) => (e.currentTarget.style.opacity = "1")}
      />
    </div>
  );
};

export default Login;