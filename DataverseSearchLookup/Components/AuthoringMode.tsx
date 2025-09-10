import {
  Button,
  Card,
  CardFooter,
  FluentProvider,
  webLightTheme,
} from "@fluentui/react-components";
import * as React from "react";
import { IInputs } from "../generated/ManifestTypes";
import {
  CheckmarkCircleColor,
  DismissCircleColor,
} from "@fluentui/react-icons";
import parse from "html-react-parser";
import { ILookUpProps } from "../types";

export type IAuthoringModeProps = Pick<
  ILookUpProps,
  "description" | "cardAppearance"
>;

export const AuthoringMode = (props: IAuthoringModeProps) => {
  const { description, cardAppearance } = props;

  return (
    <FluentProvider
      theme={webLightTheme}
      style={{ margin: 0, padding: 0, width: "100%" }}
    >
      <Card style={{ width: "100%" }} appearance={cardAppearance}>
        {parse(description)}

        <CardFooter>
          <Button icon={<CheckmarkCircleColor fontSize={16} />}>Confirm</Button>
          <Button icon={<DismissCircleColor fontSize={16} />}>Dismiss</Button>
        </CardFooter>
      </Card>
    </FluentProvider>
  );
};
