import React, { useState } from "react";
import { Stack, Text, Dropdown, PrimaryButton, DefaultButton, TextField } from "@fluentui/react";
import "./App.css";

const VSOAssistant: React.FC = () => {
  const [dataCenter, setDataCenter] = useState<string | undefined>();
  const [diversity, setDiversity] = useState<string | undefined>();
  const [spliceRack, setSpliceRack] = useState<string>("");

  const handleSubmit = () => {
    // 🔹 Here you’ll trigger your Logic App flow
    // e.g., call https://yourvsoflow.azurewebsites.net/api/vsoflow
    console.log("Submitted:", { dataCenter, diversity, spliceRack });
  };

  return (
    <div className="main-content">
      <Stack tokens={{ childrenGap: 10 }} className="vso-assistant-container">
        <Text variant="xLarge" styles={{ root: { color: "#00aaff", fontWeight: 600 } }}>
          🚀 Fiber VSO Assistant
        </Text>

        <Dropdown
          label="Select a Data Center *"
          placeholder="Select a Data Center"
          options={[
            { key: "GVX01", text: "GVX01" },
            { key: "STO31", text: "STO31" },
            { key: "DUB14", text: "DUB14" },
          ]}
          onChange={(_, option) => setDataCenter(option?.text)}
        />

        <Dropdown
          label="Select Diversity"
          placeholder="Select a Diversity Path (Optional)"
          options={[
            { key: "East", text: "East" },
            { key: "West", text: "West" },
            { key: "Y", text: "Y" },
            { key: "Z", text: "Z" },
          ]}
          onChange={(_, option) => setDiversity(option?.text)}
        />

        <TextField
          label="Splice Rack A"
          placeholder="e.g. AM111 (Optional)"
          value={spliceRack}
          onChange={(_, val) => setSpliceRack(val || "")}
        />

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton text="Submit" onClick={handleSubmit} />
          <DefaultButton text="Help" onClick={() => alert("Coming soon")} />
        </Stack>

        <Text
          variant="small"
          styles={{
            root: {
              color: "#bbb",
              marginTop: "8px",
            },
          }}
        >
          Always verify critical data before taking action.
          <br />
          Agent developed by <b>Josh Maclean</b>, supported by the <b>CIA | Network Delivery</b> team.
          <br />
          For feedback or assistance,{" "}
          <a
            href="https://teams.microsoft.com/l/chat/0/0?users=joshmaclean@microsoft.com"
            target="_blank"
            rel="noopener noreferrer"
          >
            send a message
          </a>.
        </Text>
      </Stack>
    </div>
  );
};

export default VSOAssistant;
