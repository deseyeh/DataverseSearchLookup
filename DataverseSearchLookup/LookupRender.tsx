/* eslint-disable @typescript-eslint/no-redundant-type-constituents */
/* eslint-disable @typescript-eslint/no-unnecessary-type-assertion */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-call */
import parse from "html-react-parser";
import {
  Button,
  Card,
  CardFooter,
  FluentProvider,
  makeStyles,
  Spinner,
  useId,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import {
  CheckmarkCircleColor,
  DismissCircleColor,
} from "@fluentui/react-icons";

import * as React from "react";
import { ILookUpProps } from "./types";
import {
  searchDataverse,
  getPrimaryField,
  SearchResult,
} from "./DataverseSearchHelper";
import { useEffect, useState } from "react";
import { replaceTokensWithFieldValues } from "./replaceTokensWithFieldValues";

const useStyles = makeStyles({
  root: {
    width: "100%",
    display: "grid",
    gridTemplateRows: "auto auto",
    justifyItems: "start",
    gap: "2px",
  },
});

const LookupRender = (props: ILookUpProps) => {
  const {
    context,
    fieldLogicName,
    lookupEntityName,
    lookupRecordId,
    description,
    cardAppearance,
    searchParameter,
  } = props;
  const [searchResults, setSearchResults] = useState<SearchResult | undefined>(
    undefined
  );
  const [isLoading, setIsLoading] = useState<boolean>(true);

  const [searchResultPrimaryField, setSearchResultPrimaryField] =
    useState<string>("");
  const [tokenDescription, setTokenDescription] = useState<string>(description);

  const comboboxId = useId("combobox");
  const styles = useStyles();

  useEffect(() => {
    const fetchData = async () => {
      try {
        // Resolve primary field name for entity
        const [primaryFieldName] = await getPrimaryField(
          context,
          lookupEntityName
        );
        const primaryFieldToken = `{{${fieldLogicName}.${primaryFieldName}}}`;
        // Query Dataverse for record
        const record = await searchDataverse(
          context,
          searchParameter,
          lookupEntityName
        );
        if (!record) {
          console.warn(
            "No Dataverse Search results found for:",
            searchParameter,
            "; with Primary field resolved:",
            lookupEntityName,
            primaryFieldName
          );
          setSearchResults(record);
          setSearchResultPrimaryField("");
          setTokenDescription("");
          return;
        }

        const primaryFieldValue = record.Attributes[primaryFieldName] as string;
        // Build record URL
        const appId =
          typeof Xrm !== "undefined"
            ? Xrm.Utility.getGlobalContext().getCurrentAppUrl()?.split("=")[1]
            : "5b1a454f-fe71-f011-b4cc-7c1e5278db54"; // TODO: confirm default fallback

        const orgUrl = `${window.location.protocol}//${window.location.host}`;
        const recordUrl = `${orgUrl}/main.aspx?appid=${appId}&pagetype=entityrecord&etn=${lookupEntityName}&id=${record.Id}`;
        const hyperlink = `<a href='${recordUrl}' target='_blank'>${primaryFieldValue}</a>`;

        // Replace tokens in description
        const descriptionWithUrl =
          description?.replace(primaryFieldToken, `"${hyperlink}"`) ?? "";

        const finalDescription = await replaceTokensWithFieldValues(
          descriptionWithUrl,
          context
        );

        // Update state in a single batch
        setSearchResults(record);
        setSearchResultPrimaryField(primaryFieldValue);
        setTokenDescription(finalDescription);
      } catch (error) {
        console.error("[LookupControl] fetchData failed:", error);
      } finally {
        setIsLoading(false);
      }
    };

    void fetchData();
  }, [searchParameter, lookupRecordId, lookupEntityName, description, context]);

  const handleConfirm = () => {
    if (searchResults) {
      props.onConfirm([
        {
          id: searchResults.Id,
          name: searchResultPrimaryField,
          entityType: lookupEntityName,
        },
      ]);
    }
    handleDismiss();
  };
  const handleDismiss = () => {
    props.onDismiss();
  };

  return isLoading ? (
    <Spinner size="extra-tiny" />
  ) : searchResults ? (
    <FluentProvider
      theme={webLightTheme}
      style={{ margin: "0px", padding: "0px", width: "100%" }}
    >
      <div className={styles.root}>
        <Card style={{ width: "100%" }} appearance={cardAppearance}>
          {parse(tokenDescription)}
          <CardFooter>
            <Button
              icon={<CheckmarkCircleColor fontSize={16} />}
              onClick={handleConfirm}
            >
              Confirm
            </Button>
            <Button
              icon={<DismissCircleColor fontSize={16} />}
              onClick={handleDismiss}
            >
              Dismiss
            </Button>
          </CardFooter>
        </Card>
      </div>
    </FluentProvider>
  ) : null;
};

export default LookupRender;
