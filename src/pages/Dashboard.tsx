import React from "react";
import { Stack, Text, Image } from "@fluentui/react";
import "../Theme.css";
import logo from "../assets/optical360-logo.png"; // ✅ add this

const Dashboard: React.FC = () => {
  return (
    <Stack
      horizontalAlign="center"
      verticalAlign="center"
      tokens={{ childrenGap: 20 }}
      styles={{
        root: {
          height: "100%",
          color: "#ccc",
          backgroundColor: "#111",
        },
      }}
    >
      <Image src={logo} alt="Optical 360 Logo" width={200} className="logo-img" /> {/* glowing logo */}
      <Text variant="xxLarge" styles={{ root: { color: "#50b3ff", fontWeight: 600 } }}>
        Welcome to Optical 360
      </Text>
      <Text variant="medium">Select a page from the menu to get started.</Text>
    </Stack>
  );
};

export default Dashboard;
