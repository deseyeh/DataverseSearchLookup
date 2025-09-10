/* eslint-disable @typescript-eslint/no-unsafe-argument */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
import parse from "html-react-parser";
import {
  Button,
  Card,
  CardFooter,
  FluentProvider,
  Spinner,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import {
  CheckmarkCircleColor,
  DismissCircleColor,
} from "@fluentui/react-icons";
import * as React from "react";
import { ILookUpProps } from "../types";
import {
  searchDataverse,
  getPrimaryField,
  SearchResult,
} from "../helper/DataverseSearchHelper";
import { replaceTokensWithFieldValues } from "../replaceTokensWithFieldValues";

const LookupRender: React.FC<ILookUpProps> = ({
  context,
  fieldLogicName,
  lookupEntityName,
  lookupRecordId,
  description,
  cardAppearance,
  searchParameter,
  onConfirm,
  onDismiss,
}) => {
  const [error, setError] = React.useState<string | null>(null);
  const [state, setState] = React.useState<{
    record?: SearchResult;
    primaryField: string;
    description: string;
    loading: boolean;
  }>({
    record: undefined,
    primaryField: "",
    description,
    loading: true,
  });

  React.useEffect(() => {
    const fetchData = async () => {
      try {
        const [primaryFieldName = "name"] = await getPrimaryField(
          context,
          lookupEntityName
        );

        if (!searchParameter || searchParameter.length < 2) {
          setState((prev) => ({ ...prev, loading: false }));
          setError(
            `Search parameter is either too short or invalid. Select from lookup.`
          );
          return;
        }
        const record = await searchDataverse(
          context,
          searchParameter,
          lookupEntityName
        );

        // early return if no record
        if (!record) {
          setError(
            `No match found for "${searchParameter}" on entity ${lookupEntityName}. Select from lookup.`
          );
          console.warn(
            "[LookupRender] No results found for",
            searchParameter,
            "on entity",
            lookupEntityName
          );
          setState({
            record: undefined,
            primaryField: "",
            description: "",
            loading: false,
          });
          return;
        }

        const primaryFieldValue =
          (record.Attributes[primaryFieldName] as string) ?? "(unnamed)";

        const appId =
          typeof Xrm !== "undefined"
            ? Xrm.Utility.getGlobalContext().getCurrentAppUrl()?.split("=")[1]
            : undefined;

        const recordUrl = buildRecordUrl(lookupEntityName, record.Id, appId);

        const token = `{{${fieldLogicName}.${primaryFieldName}}}`;
        const hyperlink = `<a href='${recordUrl}' target='_blank'>${primaryFieldValue}</a>`;

        const descriptionWithUrl =
          description?.replace(token, `"${hyperlink}"`) ?? "";

        const finalDescription = await replaceTokensWithFieldValues(
          descriptionWithUrl,
          context
        );

        setState({
          record,
          primaryField: primaryFieldValue,
          description: finalDescription,
          loading: false,
        });
      } catch (error) {
        console.error("[LookupRender] fetchData failed:", error);
        setError("[LookupRender] fetchData failed:");
        setState((prev) => ({ ...prev, loading: false }));
      }
    };

    void fetchData();
  }, [
    searchParameter,
    lookupRecordId,
    lookupEntityName,
    description,
    context,
    fieldLogicName,
  ]);

  const handleConfirm = () => {
    if (state.record) {
      onConfirm([
        {
          id: state.record.Id,
          name: state.primaryField,
          entityType: lookupEntityName,
        },
      ]);
    } else {
      onConfirm([]);
    }
    onDismiss();
  };

  if (state.loading) {
    return <Spinner size="tiny" label="Loading..." />;
  }
  if (error) {
    return (
      <div
        role="alert"
        style={{
          backgroundColor: "#f8d7da",
          color: "#721c24",
          padding: "2px",
          border: "1px solid #f5c6cb",
          borderRadius: "4px",
          display: "flex",
          alignItems: "center",
          gap: "8px",
        }}
      >
        <DismissCircleColor primaryFill="#721c24" />
        <span>
          <strong>Error: </strong>
          {error}
        </span>
      </div>
    );
  }
  return (
    <FluentProvider
      theme={
        context.fluentDesignLanguage?.isDarkTheme ? webDarkTheme : webLightTheme
      }
      style={{ margin: 0, padding: 0, width: "100%" }}
    >
      <Card style={{ width: "100%" }} appearance={cardAppearance}>
        {parse(state.description)}
        {
          /**show Card FOoter if there is record ID */
          state.record?.Id && (
            <CardFooter>
              <Button
                icon={<CheckmarkCircleColor fontSize={16} />}
                onClick={handleConfirm}
              >
                Confirm
              </Button>
              <Button
                icon={<DismissCircleColor fontSize={16} />}
                onClick={onDismiss}
              >
                Dismiss
              </Button>
            </CardFooter>
          )
        }
      </Card>
    </FluentProvider>
  );
};

const buildRecordUrl = (entity: string, id: string, appId?: string) => {
  const orgUrl = `${window.location.protocol}//${window.location.host}`;
  return `${orgUrl}/main.aspx?appid=${
    appId ?? ""
  }&pagetype=entityrecord&etn=${entity}&id=${id}`;
};

export default LookupRender;
