import * as React from 'react';
import { useState, useEffect } from 'react';
import { getAddress } from './getAddress';
import { getTenantInfo } from './getTenantInfo';
import { getContactInfo } from './getContactInfo';
import { IInputs } from './generated/ManifestTypes';
import './css/PCF2.css';
import { Label, Spinner, SpinnerSize } from '@fluentui/react';

export interface ITenantAndPropertyInfoProps {
  id: string | null;
  context: ComponentFramework.Context<IInputs>;
}

export const TenantAndPropertyInfo: React.FC<ITenantAndPropertyInfoProps> = ({ id, context }) => {

  // State hook to store the address of the property
  const [address, setAddress] = useState<string>('');

  // State hooks to store tenant name
  const [tenantName, setTenantName] = useState<string | null>(null);

  // State hooks to store tenant contact phone
  const [mobilePhone, setMobilePhone] = useState<string | null>(null);

  // State hooks to store tenant contact email
  const [emailAddress, setEmailAddress] = useState<string | null>(null);

  // State hook to store loading state
  const [loading, setLoading] = useState<boolean>(true);

  // State hook to store error message
  const [error, setError] = useState<string | null>(null);

  // Fetch data when the component mounts
  useEffect(() => {
    // Check if id is null and display a message instead of fetching data
    if (!id) {
      setError("No information to display. Please select a property.");
      setLoading(false);
      return;
    }

    // Fetch data from the Web API
    const fetchData = async () => {
      try {
        // Fetch address and property ID from the property related to the maintenance request
        const { address, tenantId } = await getAddress(id, context);
        setAddress(address);

        // Fetch tenant name of the tenant related to the property, if tenantId is not null
        const { tenantName } = await getTenantInfo(tenantId, context);
        setTenantName(tenantName);

        // Fetch contact info related to the tenant, if tenantId is not null 
        const { mobilePhone, emailAddress } = await getContactInfo(tenantId, context);
        setMobilePhone(mobilePhone);
        setEmailAddress(emailAddress);

      } catch (error) {
        console.error("Error fetching data:", error);
        setError("An error occurred while fetching data.");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [id, context]);

  // Render the component based on the loading state
  if (loading) {
    return (
      <div className="loading-container">
        <Spinner label="Loading data..." size={SpinnerSize.large} />
      </div>
    );
  }

  // Render the component based on the error message
  if (error) {
    return (
      <div className="no-property-card">
        <Label className="no-property-message">{error}</Label>
      </div>
    );
  }

  return (
    <div className="outer-container">
      <Label className="section-label">Address:</Label>
      <div className="address-container">{address}</div>
      
      <Label className="section-label">Tenant:</Label>
      <div className="tenantInfo-container">
        <div className="tenant-name">
          {tenantName ? tenantName : "This property is currently unoccupied"}
        </div>
        
        {tenantName && (
          <>
            <div className="separator">
              <svg className="wavy-line" width="100%" height="10px" viewBox="0 0 200 10" preserveAspectRatio="none">
                <defs>
                  <pattern id="seamlessWavePattern" x="0" y="0" width="10" height="10" patternUnits="userSpaceOnUse">
                    <path d="M0,5 Q2.5,0 5,5 T10,5" fill="transparent" stroke="#555555" strokeWidth="0.8" />
                  </pattern>
                </defs>
                <rect width="100%" height="10" fill="url(#seamlessWavePattern)" />
              </svg>
            </div>

            <div className="tenant-details">
              <div className="tenant-detail">
                <span className="detail-label">Mobile Phone:</span>
                <span className="detail-text">{mobilePhone || "No mobile phone available"}</span>
              </div>
              
              <div className="tenant-detail">
                <span className="detail-label">Email Address:</span>
                <span className="detail-text">{emailAddress || "No email address available"}</span>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );    
};
