import { IInputs } from "./generated/ManifestTypes";

/**
 * Fetches the address and tenant ID of the property record.
 * @param propertyId The ID of the property record.
 * @param context The PCF context object.
 * @returns The address and tenant ID of the property record.
 */
export const getAddress = async ( propertyId: string, context: ComponentFramework.Context<IInputs> ): Promise<{ address: string; tenantId: string | null}> => {

  // The logical name of the entity
  const entityLogicalName = "sam_property";

  // The fields to fetch from the property record
  const fileFieldName = `?$filter=sam_propertyid eq '${propertyId}'&$select=sam_address,_sam_tenant_value`;

  try {

    // Fetch the property record
    const response = await context.webAPI.retrieveMultipleRecords(entityLogicalName, fileFieldName);

    // If a record is found, return the address and tenant ID
    if (response.entities.length > 0) {
      // Get the first record
      const record = response.entities[0];
      // Get the address and tenant ID
      const address = record.sam_address;
      const tenantId = record._sam_tenant_value;
      return { address, tenantId};
    } else {
      throw new Error("No maintenance inspection record found with the given ID.");
    }
  } catch (error) {
    console.error("Error fetching maintenance inspection record:", error);
    throw error;
  }
};
