import { IInputs } from "./generated/ManifestTypes";

/**
 * Fetches the contact information of the contact record.
 * @param tenantId The ID of the tenant record.
 * @param context The PCF context object.
 * @returns The mobile phone and email address of the contact record.
 */
export const getContactInfo = async ( tenantId: string | null, context: ComponentFramework.Context<IInputs>): Promise<{ mobilePhone: string | null; emailAddress: string | null }> => {
  
  // If the tenant ID is null, return mobile phone and email address as null
  if (!tenantId) {
    return { mobilePhone: null, emailAddress: null };
  }

  // The logical name of the entity
  const entityLogicalName = "contact";

  // The fields to fetch from the contact record
  const fileFieldName = `?$filter=_sam_tenant_value eq '${tenantId}'&$select=mobilephone,emailaddress1`;

  try {

    // Fetch the contact record
    const response = await context.webAPI.retrieveMultipleRecords(entityLogicalName, fileFieldName);

    // If a record is found, return the mobile phone and email address
    if (response.entities.length > 0) {
      // Get the first record
      const record = response.entities[0];
      // Get the mobile phone and email address
      const mobilePhone = record.mobilephone;
      const emailAddress = record.emailaddress1;
      return { mobilePhone, emailAddress };
    } else {
      throw new Error("No maintenance inspection record found with the given ID.");
    }
  } catch (error) {
    console.error("Error fetching contact record:", error);
    throw error;
  }
};
