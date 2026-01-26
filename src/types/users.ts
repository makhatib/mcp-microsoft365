/**
 * Users Types
 */

export interface User {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
}

export interface UserSummary {
  id: string;
  displayName: string;
  email?: string;
  jobTitle?: string;
}
