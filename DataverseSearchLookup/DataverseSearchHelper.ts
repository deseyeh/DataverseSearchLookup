import { IInputs } from "./generated/ManifestTypes";

export interface SearchResult {
    Id: string;
    EntityName: string;
    Attributes: Record<string, unknown>;
    Score: number;
    Highlights?: Record<string, string[]>;
}

export interface SearchResponse {
    Error: unknown;
    Value: SearchResult[];
    Count: number;
}

/**
 * Get default columns for an entity (primary name field)
 */
export async function getPrimaryField(context: ComponentFramework.Context<IInputs>, entityName: string): Promise<string[]> {
    try {
        const response = await context.webAPI.retrieveRecord(
            "EntityDefinition",
            `(LogicalName='${entityName}')`,
            "?$select=PrimaryNameAttribute"
      ); 

        return [response.PrimaryNameAttribute as string];
    } catch {
        // Fallback to common primary name fields
        const commonFields: Record<string, string> = {
            'contact': 'fullname',
            'account': 'name',
      }; 

        return [commonFields[entityName] || 'name'];
    }
}

/**
 * Simple async search using modern searchquery endpoint
 * @param context PCF context
 * @param searchTerm What to search for
 * @param entityName Entity to search (e.g., 'contact', 'account')
 * @param selectColumns Columns to return (default: primary name field)
 * @param top Max results (default: 10)
 */
export async function searchDataverse(
    context: ComponentFramework.Context<IInputs>,
    searchTerm: string,
    entityName: string,
    selectColumns: string[] = [],
    top = 5
): Promise<SearchResult | undefined> {
    if (!searchTerm || searchTerm.length < 2) {
        return undefined;
    }

    try {
        // Get the organization URL - use window.location for PCF controls
      const orgUrl = `${window.location.protocol}//${window.location.host}`; 
        
        // Build entity configuration
        const columns = selectColumns.length > 0 ? selectColumns : await getPrimaryField(context, entityName);
        const entityConfig = {
            Name: entityName,
            // SelectColumns: columns,
            // SearchColumns: columns,
            // Filter: "statecode eq 0" // Only active records
        }; 

        // Build search request body
        const searchBody = {
            search: searchTerm,
            top: top,
            entities: JSON.stringify([entityConfig]),
            count: true
        }; 

        // Make the search call using fetch
        const response = await fetch(`${orgUrl}/api/data/v9.2/searchquery`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json'
            },
            body: JSON.stringify(searchBody)
        });

        if (!response.ok) {
            throw new Error(`Search failed: ${response.status} ${response.statusText}`);
        }

        const result = await response.json() as { response: string };
        // Parse the nested response
        const searchResponse = JSON.parse(result.response) as SearchResponse;
        
        // Return the first result instead of the entire array
        return searchResponse.Value?.[0];

    } catch (error) {
        console.error('Search failed:', error);
        return undefined;
    }
}

