/* eslint-disable @typescript-eslint/no-unsafe-argument */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */

import { IInputs } from "./generated/ManifestTypes";
const FORMATTED_VALUE_SUFFIX = "@OData.Community.Display.V1.FormattedValue";

export async function replaceTokensWithFieldValues(
  template: string,
  context: ComponentFramework.Context<IInputs>,
  lookupOverride?: { id: string; entity: string }
): Promise<string> {
  template = resolveBuiltinTokens(template);

  const tokens = [...template.matchAll(/{{(.*?)}}/g)].map((m) => m[1].trim());
  const uniqueTokens = [...new Set(tokens)];

  const rawTokenSet = new Set(
    uniqueTokens
      .map(t => {
        const fieldName = t.split(".").slice(-1)[0];
        if (/^_.*_value$/.test(fieldName)) {
          return fieldName;
        }
        return null;
      })
      .filter(Boolean)
  );

  const defaultEntityId = (context.mode as any).contextInfo.entityId;
  const defaultEntityName = (context.mode as any).contextInfo.entityTypeName;

  const resolvedValues: Record<string, string> = {};
  const lookupMap: Record<string, Set<string>> = {};

  for (const token of uniqueTokens) {
    const parts = token.split(".");
    const lookupPath = parts.slice(0, -1).join(".");
    const finalField = parts[parts.length - 1];
    const key = lookupPath || "_self";

    if (!lookupMap[key]) {
      lookupMap[key] = new Set();
    }
    lookupMap[key].add(finalField);
  }

  const cache: Record<string, any> = {};

  for (const [lookupPath, fields] of Object.entries(lookupMap)) {
    try {
      const pathParts = lookupPath !== "_self" ? lookupPath.split(".") : [];

      let currentEntity = defaultEntityName;
      let currentId = defaultEntityId;

      let remainingPath = pathParts;

      // Determine if override applies
      let overrideApplied = false;
      let firstLookupField = "";

      if (lookupOverride && pathParts.length > 0) {
        firstLookupField = pathParts[0];
        const actualLookupEntity = await getLookupEntityName(context, currentEntity, currentId, firstLookupField);
        if (actualLookupEntity && actualLookupEntity.toLowerCase() === lookupOverride.entity.toLowerCase()) {
          currentEntity = lookupOverride.entity;
          currentId = lookupOverride.id;
          remainingPath = pathParts.slice(1);
          overrideApplied = true;
        }
      }

      // Traverse remaining path
      for (const part of remainingPath) {
        const cacheKey = `${currentEntity}:${currentId}:${part}`;
        if (cache[cacheKey]) {
          ({ id: currentId, entity: currentEntity } = cache[cacheKey]);
          continue;
        }

        const lookupField = `_${part}_value`;
        const lookupLogicalNameField = `${lookupField}@Microsoft.Dynamics.CRM.lookuplogicalname`;

        const record = await context.webAPI.retrieveRecord(
          currentEntity,
          currentId,
          `?$select=${lookupField}`
        );

        const nextId = record[lookupField];
        const nextEntity = record[lookupLogicalNameField];

        if (!nextId || !nextEntity) {
          throw new Error(`Failed to resolve lookup field "${part}" on entity "${currentEntity}" with id "${currentId}"`);
        }

        currentId = nextId;
        currentEntity = nextEntity;

        cache[cacheKey] = { id: currentId, entity: currentEntity };
      }

      // Fetch final fields
      const fieldList = [...fields].map((f) =>
        /^_.*_value$/.test(f) ? f.substring(1, f.length - 6) : f
      ).join(",");

      const recordRaw = await context.webAPI.retrieveRecord(
        currentEntity,
        currentId,
        `?$select=${fieldList}`
      );

      // Resolve fields
      for (const field of fields) {
        const token = lookupPath === "_self" ? field : `${lookupPath}.${field}`;
        const isRaw = rawTokenSet.has(field);

        let actualFieldName = field;
        if (isRaw && /^_.*_value$/.test(field)) {
          actualFieldName = field.substring(1, field.length - 6);
        }

        if (isRaw) {
          resolvedValues[token] = recordRaw[actualFieldName] ?? "";
        } else {
          const formattedKey = `${actualFieldName}${FORMATTED_VALUE_SUFFIX}`;
          resolvedValues[token] = recordRaw[formattedKey] ?? recordRaw[actualFieldName] ?? "";
        }
      }
 

    } catch (err) {
      console.error(`Error resolving path "${lookupPath}":`, err);
      for (const field of fields) {
        const token = lookupPath === "_self" ? field : `${lookupPath}.${field}`;
        resolvedValues[token] = "";
      }
    }
  }

  const result = template.replace(/{{(.*?)}}/g, (_, key) => {
    const value = resolvedValues[key.trim()];
    return value !== undefined && value !== null ? String(value) : "";
  });

  return result;
}

async function getLookupEntityName(
  context: ComponentFramework.Context<IInputs>,
  baseEntity: string,
  baseId: string,
  lookupField: string
): Promise<string | null> {
  try {
    const lookupFieldName = `_${lookupField}_value`;
    const logicalNameField = `${lookupFieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`;
    const record = await context.webAPI.retrieveRecord(
      baseEntity,
      baseId,
      `?$select=${lookupFieldName}`
    );
    const value = record[logicalNameField];
    return typeof value === "string" ? value : null;
  } catch (e) {
    console.error("Error fetching lookup entity name:", e);
    return null;
  }
}

export function resolveBuiltinTokens(template: string): string {
  const builtins: Record<string, string> = {
    "getClientUrl()": Xrm.Utility.getGlobalContext().getClientUrl(),
    "userId()": Xrm.Utility.getGlobalContext().userSettings.userId,
    "userName()": Xrm.Utility.getGlobalContext().userSettings.userName,
    "today()": new Date().toISOString().split("T")[0],
    "now()": new Date().toISOString(),
  };
  return template.replace(/{{(.*?)}}/g, (_, token) => {
    const trimmed = token.trim();
    return Object.prototype.hasOwnProperty.call(builtins, trimmed)
      ? builtins[trimmed]
      : `{{${trimmed}}}`;
  });
  
}
