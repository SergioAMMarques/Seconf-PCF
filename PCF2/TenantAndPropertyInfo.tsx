import * as React from 'react';
import { useState, useEffect } from 'react';
import { getAddress } from './getAddress';
import { getTenantInfo } from './getTenantInfo';
import { getContactInfo } from './getContactInfo';
import { IInputs } from './generated/ManifestTypes';
import './css/PCF2.css'
import { Label } from '@fluentui/react';

export interface ITenantAndPropertyInfoProps {
  id: string;
  context: ComponentFramework.Context<IInputs>;
}

export const TenantAndPropertyInfo: React.FC<ITenantAndPropertyInfoProps> = ({ id, context }) => {

  // State to store the address 
  const [address, setAddress] = useState<string>('');

  // State to store the tenant name
  const [tenantName, setTenantName] = useState<string | null>(null);

  // State to store the mobile phone
  const [mobilePhone, setMobilePhone] = useState<string | null>(null);

  // State to store the email address
  const [emailAddress, setEmailAddress] = useState<string | null>(null);

  // State to store the loading status
  const [loading, setLoading] = useState<boolean>(true);

  // State to store the error message
  const [error, setError] = useState<string | null>(null);

  // Fetch data when the component mounts
  useEffect(() => {
    const fetchData = async () => {
      try {
        // Fetch address and property ID from the property related to the maintenance request
        const { address, tenantId } = await getAddress(id, context);
        setAddress(address);

        // Fetch tenant name of the tenant realeted to the property, if tenantId is not null
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

  if (loading) {
    return <div>Loading...</div>;
  }

  if (error) {
    return <div>{error}</div>;
  }

  return (
    <>
      <div className="outer-container">
        <Label className="section-label">Address:</Label>
        <div className="address-container">{address}</div>
        
        <Label className="section-label">Tenant:</Label>
        <div className="tenantInfo-container">
          <div className="tenant-name">{tenantName || "No tenant assigned"}</div>
          
          <div className="separator">
            <svg className="wavy-line" width="100%" height="10px" viewBox="0 0 200 10" preserveAspectRatio="none">
              <path d="M0,5 Q2.5,0 5,5 T10,5 T15,5 T20,5 T25,5 T30,5 T35,5 T40,5 T45,5 T50,5 T55,5 T60,5 T65,5 T70,5 T75,5 T80,5 T85,5 T90,5 T95,5 T100,5 T105,5 T110,5 T115,5 T120,5 T125,5 T130,5 T135,5 T140,5 T145,5 T150,5 T155,5 T160,5 T165,5 T170,5 T175,5 T180,5 T185,5 T190,5 T195,5 T200,5"
                    fill="transparent" stroke="#555555" strokeWidth="1.2"/>
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
        </div>
      </div>
    </>
  );    
};
