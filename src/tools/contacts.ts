/**
 * Contacts Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';

// === Types ===

interface Contact {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  emailAddresses?: Array<{ address: string; name?: string }>;
  businessPhones?: string[];
  mobilePhone?: string;
  companyName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessAddress?: {
    street?: string;
    city?: string;
    state?: string;
    countryOrRegion?: string;
    postalCode?: string;
  };
}

interface ContactSummary {
  id: string;
  displayName: string;
  email?: string;
  phone?: string;
  company?: string;
}

// === Schemas ===

const contactsListSchema = z.object({
  user: z.string().optional(),
  top: z.number().min(1).max(100).default(20),
  filter: z.string().optional(),
});

const contactsGetSchema = z.object({
  user: z.string().optional(),
  contactId: z.string(),
});

const contactsCreateSchema = z.object({
  user: z.string().optional(),
  givenName: z.string(),
  surname: z.string().optional(),
  email: z.string().optional(),
  phone: z.string().optional(),
  mobilePhone: z.string().optional(),
  companyName: z.string().optional(),
  jobTitle: z.string().optional(),
  department: z.string().optional(),
});

const contactsUpdateSchema = z.object({
  user: z.string().optional(),
  contactId: z.string(),
  givenName: z.string().optional(),
  surname: z.string().optional(),
  email: z.string().optional(),
  phone: z.string().optional(),
  mobilePhone: z.string().optional(),
  companyName: z.string().optional(),
  jobTitle: z.string().optional(),
  department: z.string().optional(),
});

const contactsSearchSchema = z.object({
  user: z.string().optional(),
  query: z.string(),
  top: z.number().min(1).max(50).default(10),
});

const contactsDeleteSchema = z.object({
  user: z.string().optional(),
  contactId: z.string(),
});

// === Tool Definitions ===

export const contactsTools: Tool[] = [
  {
    name: 'm365_contacts_list',
    description: 'List contacts from user\'s address book',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        top: { type: 'number', description: 'Number of contacts to return (1-100)', default: 20 },
        filter: { type: 'string', description: 'OData filter (e.g., "companyName eq \'Contoso\'")' },
      },
    },
  },
  {
    name: 'm365_contacts_get',
    description: 'Get detailed information about a specific contact',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        contactId: { type: 'string', description: 'Contact ID' },
      },
      required: ['contactId'],
    },
  },
  {
    name: 'm365_contacts_create',
    description: 'Create a new contact',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        givenName: { type: 'string', description: 'First name' },
        surname: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        phone: { type: 'string', description: 'Business phone' },
        mobilePhone: { type: 'string', description: 'Mobile phone' },
        companyName: { type: 'string', description: 'Company name' },
        jobTitle: { type: 'string', description: 'Job title' },
        department: { type: 'string', description: 'Department' },
      },
      required: ['givenName'],
    },
  },
  {
    name: 'm365_contacts_update',
    description: 'Update an existing contact',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        contactId: { type: 'string', description: 'Contact ID to update' },
        givenName: { type: 'string', description: 'First name' },
        surname: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        phone: { type: 'string', description: 'Business phone' },
        mobilePhone: { type: 'string', description: 'Mobile phone' },
        companyName: { type: 'string', description: 'Company name' },
        jobTitle: { type: 'string', description: 'Job title' },
        department: { type: 'string', description: 'Department' },
      },
      required: ['contactId'],
    },
  },
  {
    name: 'm365_contacts_search',
    description: 'Search contacts by name or email',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        query: { type: 'string', description: 'Search query (name or email)' },
        top: { type: 'number', description: 'Number of results', default: 10 },
      },
      required: ['query'],
    },
  },
  {
    name: 'm365_contacts_delete',
    description: 'Delete a contact',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        contactId: { type: 'string', description: 'Contact ID to delete' },
      },
      required: ['contactId'],
    },
  },
];

// === Handlers ===

async function contactsList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = contactsListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_contacts_list', { top: input.top });

  const response = await graphClient.get<GraphListResponse<Contact>>(
    `/users/${user}/contacts`,
    {
      $top: input.top,
      $select: 'id,displayName,emailAddresses,businessPhones,mobilePhone,companyName',
      $orderby: 'displayName',
      $filter: input.filter,
    }
  );

  const result: ContactSummary[] = response.value.map(c => ({
    id: c.id,
    displayName: c.displayName,
    email: c.emailAddresses?.[0]?.address,
    phone: c.businessPhones?.[0] || c.mobilePhone,
    company: c.companyName,
  }));

  logToolResult('m365_contacts_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function contactsGet(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = contactsGetSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_contacts_get', { contactId: input.contactId });

  const contact = await graphClient.get<Contact>(
    `/users/${user}/contacts/${input.contactId}`
  );

  const result = {
    id: contact.id,
    displayName: contact.displayName,
    givenName: contact.givenName,
    surname: contact.surname,
    emails: contact.emailAddresses,
    businessPhones: contact.businessPhones,
    mobilePhone: contact.mobilePhone,
    companyName: contact.companyName,
    jobTitle: contact.jobTitle,
    department: contact.department,
    officeLocation: contact.officeLocation,
    businessAddress: contact.businessAddress,
  };

  logToolResult('m365_contacts_get', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function contactsCreate(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = contactsCreateSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_contacts_create', { givenName: input.givenName });

  const contactData: Record<string, unknown> = {
    givenName: input.givenName,
  };

  if (input.surname) contactData.surname = input.surname;
  if (input.email) {
    contactData.emailAddresses = [{ address: input.email }];
  }
  if (input.phone) contactData.businessPhones = [input.phone];
  if (input.mobilePhone) contactData.mobilePhone = input.mobilePhone;
  if (input.companyName) contactData.companyName = input.companyName;
  if (input.jobTitle) contactData.jobTitle = input.jobTitle;
  if (input.department) contactData.department = input.department;

  const created = await graphClient.post<Contact>(
    `/users/${user}/contacts`,
    contactData
  );

  logToolResult('m365_contacts_create', true, Date.now() - start);
  return JSON.stringify({
    success: true,
    id: created.id,
    displayName: created.displayName,
  }, null, 2);
}

async function contactsUpdate(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = contactsUpdateSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_contacts_update', { contactId: input.contactId });

  const updates: Record<string, unknown> = {};

  if (input.givenName) updates.givenName = input.givenName;
  if (input.surname) updates.surname = input.surname;
  if (input.email) {
    updates.emailAddresses = [{ address: input.email }];
  }
  if (input.phone) updates.businessPhones = [input.phone];
  if (input.mobilePhone) updates.mobilePhone = input.mobilePhone;
  if (input.companyName) updates.companyName = input.companyName;
  if (input.jobTitle) updates.jobTitle = input.jobTitle;
  if (input.department) updates.department = input.department;

  await graphClient.patch(
    `/users/${user}/contacts/${input.contactId}`,
    updates
  );

  logToolResult('m365_contacts_update', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Contact updated' });
}

async function contactsSearch(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = contactsSearchSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_contacts_search', { query: input.query });

  // Search by displayName or email containing the query
  const filter = `startswith(displayName,'${input.query}') or startswith(givenName,'${input.query}') or startswith(surname,'${input.query}')`;

  const response = await graphClient.get<GraphListResponse<Contact>>(
    `/users/${user}/contacts`,
    {
      $top: input.top,
      $filter: filter,
      $select: 'id,displayName,emailAddresses,businessPhones,mobilePhone,companyName',
    }
  );

  const result: ContactSummary[] = response.value.map(c => ({
    id: c.id,
    displayName: c.displayName,
    email: c.emailAddresses?.[0]?.address,
    phone: c.businessPhones?.[0] || c.mobilePhone,
    company: c.companyName,
  }));

  logToolResult('m365_contacts_search', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function contactsDelete(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = contactsDeleteSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_contacts_delete', { contactId: input.contactId });

  await graphClient.delete(`/users/${user}/contacts/${input.contactId}`);

  logToolResult('m365_contacts_delete', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Contact deleted' });
}

// === Export ===

export const contactsHandlers: Record<string, ToolHandler> = {
  'm365_contacts_list': contactsList,
  'm365_contacts_get': contactsGet,
  'm365_contacts_create': contactsCreate,
  'm365_contacts_update': contactsUpdate,
  'm365_contacts_search': contactsSearch,
  'm365_contacts_delete': contactsDelete,
};
