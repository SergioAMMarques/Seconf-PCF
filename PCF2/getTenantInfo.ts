import { IInputs } from "./generated/ManifestTypes";

/**
 * Fetches the tenant name of the tenant record.
 * @param tenantId The ID of the tenant record.
 * @param context the PCF context object.
 * @returns The tenant name of the tenant record.
 */
export const getTenantInfo = async ( tenantId: string | null, context: ComponentFramework.Context<IInputs> ): Promise<{ tenantName: string | null}> => {

  // If the tenant ID is null, return full name as null
  if (!tenantId) {
    return { tenantName: null };
  }

  // The logical name of the entity
  const entityLogicalName = "sam_tenant";

  // The fields to fetch from the tenant record
  const fileFieldName = `?$filter=sam_tenantid eq '${tenantId}'&$select=sam_fullname`;

  try {

    // Fetch the tenant record
    const response = await context.webAPI.retrieveMultipleRecords(entityLogicalName, fileFieldName);

    // If a record is found, return the tenant name
    if (response.entities.length > 0) {
      // Get the first record
      const record = response.entities[0];
      // Get the tenant name
      const tenantName = record.sam_fullname;
      return { tenantName };
    } else {
      throw new Error("No maintenance inspection record found with the given ID.");
    }
  } catch (error) {
    console.error("Error fetching property record:", error);
    throw error;
  }
};
