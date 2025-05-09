// "use client" must remain at the top for React hooks to work
"use client";

import React, { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Checkbox } from "@/components/ui/checkbox";
import { Card, CardContent } from "@/components/ui/card";
import { useMsal } from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";

function AuthButtons() {
  const { instance, accounts } = useMsal();

  const signIn = () => {
    instance.loginPopup({
      scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],
    });
  };

  const signOut = () => {
    instance.logoutPopup();
  };

  return accounts.length > 0 ? (
    <div className="flex items-center justify-between w-full">
      <p>Signed in as: {accounts[0].username}</p>
      <Button onClick={signOut}>Sign out</Button>
    </div>
  ) : (
    <Button onClick={signIn}>Sign in with Microsoft</Button>
  );
}

interface SiteItem {
  name: string;
  type: "List" | "Document Library" | "Page";
  selected: boolean;
}

interface ListColumn {
  Title: string;
  FieldTypeKind: number;
  Required: boolean;
  Hidden: boolean;
  ReadOnlyField: boolean;
  __metadata?: { type: string };
}

interface ListField {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Required: boolean;
  Hidden: boolean;
  ReadOnlyField: boolean;
  FieldTypeKind: number;
  Choices?: string[];
  DefaultValue?: string;
  LookupList?: string;
  LookupField?: string;
  __metadata?: {
    type: string;
  };
}

interface ListView {
  Title: string;
  ViewFields: {
    results: string[];
  };
  ViewQuery: string;
  RowLimit: number;
  Paged: boolean;
  DefaultView: boolean;
  __metadata: {
    type: string;
    uri?: string;
    id?: string;
  };
  Id?: string;
}

interface ListItemBasic {
  Id: number;
  Title: string;
  __metadata?: {
    type: string;
    uri: string;
  };
}

interface ListItemFull {
  Id: number;
  Title: string;
  [key: string]: any;
}

interface SharePointResponse<T> {
  d: {
    results: T[];
  };
}

interface GraphListItem {
  name: string;
  list?: {
    template: string;
  };
}

interface GraphPageItem {
  name: string;
}

// Helper functions for list operations
async function deleteListIfExists(hostname: string, path: string, listName: string, accessToken: string) {
  try {
    const checkListRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (checkListRes.ok) {
      const deleteListRes = await fetch(
        `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')`,
        {
          method: "DELETE",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
          },
        }
      );

      if (deleteListRes.ok) {
        return true;
      }
    }
    return false;
  } catch (error) {
    return false;
  }
}

async function createList(hostname: string, path: string, listName: string, baseTemplate: number, accessToken: string) {
  const createRes = await fetch(
    `https://${hostname}${path}/_api/web/lists`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      body: JSON.stringify({
        __metadata: { type: "SP.List" },
        Title: listName,
        BaseTemplate: baseTemplate,
      }),
    }
  );

  if (!createRes.ok) {
    throw new Error(`Failed to create list: ${await createRes.text()}`);
  }

  return await createRes.json();
}

async function createListField(hostname: string, path: string, listName: string, field: ListField, accessToken: string) {
  try {
    // Use AddFieldAsXml for Choice fields
    if (field.TypeAsString === 'Choice' && field.Choices) {
      // Support both array and object-with-results for Choices
      const choicesArray = Array.isArray(field.Choices)
        ? field.Choices
        : (typeof field.Choices === 'object' && field.Choices && 'results' in field.Choices && Array.isArray((field.Choices as any).results)
            ? (field.Choices as any).results
            : []);
      const choicesXml = choicesArray.map((choice: string) => `<CHOICE>${choice}</CHOICE>`).join('');
      const fieldXml = `
        <Field Type="Choice" Name="${field.InternalName}" DisplayName="${field.Title}" Format="Dropdown">
          <CHOICES>${choicesXml}</CHOICES>
        </Field>
      `;
      const xmlPayload = {
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: fieldXml,
          Options: 0
        }
      };
      const createFieldRes = await fetch(
        `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/fields/CreateFieldAsXml`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          body: JSON.stringify(xmlPayload),
        }
      );
      if (!createFieldRes.ok) {
        throw new Error(`Failed to create field: ${await createFieldRes.text()}`);
      }
      return;
    }
    // Fallback to REST for other field types
    let fieldPayload: any = {
      __metadata: { type: "SP.Field" },
      Title: field.Title,
      FieldTypeKind: field.FieldTypeKind,
      Required: field.Required,
    };
    switch (field.TypeAsString) {
      case 'User':
        fieldPayload = {
          ...fieldPayload,
          __metadata: { type: "SP.FieldUser" },
          SelectionMode: 1 // 1 for single user, 2 for multiple users
        };
        break;
      case 'DateTime':
        fieldPayload = {
          ...fieldPayload,
          __metadata: { type: "SP.FieldDateTime" },
          DisplayFormat: 0, // 0 for date only, 1 for date and time
          DateTimeCalendarType: 0, // 0 for Gregorian
          FriendlyDisplayFormat: 0 // 0 for default
        };
        break;
      case 'Lookup':
        fieldPayload = {
          ...fieldPayload,
          __metadata: { type: "SP.FieldLookup" },
          LookupList: field.LookupList,
          LookupField: field.LookupField
        };
        break;
    }
    if (field.DefaultValue) {
      fieldPayload.DefaultValue = field.DefaultValue;
    }
    const createFieldRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/fields`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(fieldPayload),
      }
    );
    if (!createFieldRes.ok) {
      throw new Error(`Failed to create field: ${await createFieldRes.text()}`);
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Failed to create field: ${errorMessage}`);
  }
}

async function createListView(
  hostname: string,
  path: string,
  listName: string,
  view: ListView,
  accessToken: string,
  setLogs: (updater: (prev: string[]) => string[]) => void,
  customFormatter?: string
) {
  try {
    // Create view first
    const viewPayload = {
      __metadata: { type: "SP.View" },
      Title: view.Title,
      PersonalView: false,
      ViewQuery: view.ViewQuery || "",
      RowLimit: view.RowLimit || 30,
      Paged: view.Paged !== false,
      DefaultView: view.DefaultView || false
    };
    const createViewRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(viewPayload),
      }
    );
    if (!createViewRes.ok) {
      throw new Error(`Failed to create view: ${await createViewRes.text()}`);
    }
    const viewData = await createViewRes.json();
    const viewId = viewData.d.Id;
    // Add each field to the view
    const fields = view.ViewFields?.results || [];
    for (const field of fields) {
      try {
        const addFieldRes = await fetch(
          `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views('${viewId}')/ViewFields/AddViewField('${field}')`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose"
            }
          }
        );
        if (!addFieldRes.ok) {
          throw new Error(`Failed to add field to view: ${await addFieldRes.text()}`);
        }
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev: string[]) => [...prev, `⚠️ Failed to add field ${field} to view: ${errorMessage}`]);
      }
    }
    // PATCH CustomFormatter if provided
    if (customFormatter) {
      const patchRes = await fetch(
        `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views('${viewId}')`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-HTTP-Method": "MERGE"
          },
          body: JSON.stringify({ CustomFormatter: customFormatter })
        }
      );
      if (!patchRes.ok) {
        const errorText = await patchRes.text();
        setLogs((prev: string[]) => [...prev, `⚠️ Failed to patch CustomFormatter: ${errorText}`]);
      } else {
        setLogs((prev: string[]) => [...prev, `✅ Patched CustomFormatter for view: ${view.Title}`]);
      }
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Failed to create view: ${errorMessage}`);
  }
}

async function createListItem(hostname: string, path: string, listName: string, item: ListItemBasic, accessToken: string) {
  try {
    const cleanItem = Object.entries(item)
      .filter(([key, value]) => !key.startsWith('_') && !key.startsWith('__') && value !== undefined && value !== null)
      .reduce((acc, [key, value]) => {
        acc[key] = value;
        return acc;
      }, {} as Record<string, any>);

    const itemPayload = {
      __metadata: { type: "SP.ListItem" },
      ...cleanItem
    };

    const createItemRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/items`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(itemPayload),
      }
    );

    if (!createItemRes.ok) {
      throw new Error(`Failed to create item: ${await createItemRes.text()}`);
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Failed to create item: ${errorMessage}`);
  }
}

// Main cloning function
async function cloneList(
  sourceHostname: string,
  sourcePath: string,
  targetHostname: string,
  targetPath: string,
  listName: string,
  accessToken: string,
  setLogs: (updater: (prev: string[]) => string[]) => void
) {
  try {
    // Delete existing list if it exists
    const wasDeleted = await deleteListIfExists(targetHostname, targetPath, listName, accessToken);
    if (wasDeleted) {
      setLogs((prev) => [...prev, `Deleted existing list: ${listName}`]);
    }

    // Get list details from source
    const encodedListTitle = encodeURIComponent(listName);
    const restListDetailsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!restListDetailsRes.ok) {
      throw new Error(`Failed to get list details: ${await restListDetailsRes.text()}`);
    }

    const restListDetails = await restListDetailsRes.json();
    if (!restListDetails?.d?.BaseTemplate) {
      throw new Error('BaseTemplate not found in list details');
    }

    // Create the list
    await createList(targetHostname, targetPath, listName, restListDetails.d.BaseTemplate, accessToken);
    setLogs((prev) => [...prev, `✅ Created list: ${listName}`]);

    // Get and create fields
    const fieldsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/fields`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!fieldsRes.ok) {
      throw new Error(`Failed to fetch fields: ${await fieldsRes.text()}`);
    }

    const fieldsData = await fieldsRes.json() as SharePointResponse<ListField>;
    const builtInFields = ['ContentType', 'Attachments', 'Title', 'ID', 'Created', 'Modified', 'Author', 'Editor'];
    const fields = fieldsData.d.results.filter(field => 
      !field.Hidden && 
      !field.ReadOnlyField && 
      !builtInFields.includes(field.InternalName)
    );

    // Create each field
    for (const field of fields) {
      try {
        await createListField(targetHostname, targetPath, listName, field, accessToken);
        setLogs((prev) => [...prev, `✅ Created field: ${field.Title}`]);
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `⚠️ Failed to create field ${field.Title}: ${errorMessage}`]);
      }
    }

    // Get and create views
    const viewsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/views`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!viewsRes.ok) {
      throw new Error(`Failed to fetch views: ${await viewsRes.text()}`);
    }

    const viewsData = await viewsRes.json() as SharePointResponse<ListView>;
    const views = viewsData.d.results;

    // Create each view
    for (const view of views) {
      try {
        // Fetch CustomFormatter for this view
        let customFormatter: string | undefined = undefined;
        const viewDetailsRes = await fetch(
          `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/views('${view.Id}')`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
            },
          }
        );
        if (viewDetailsRes.ok) {
          const viewDetails = await viewDetailsRes.json();
          customFormatter = viewDetails.d.CustomFormatter;
        }
        await createListView(targetHostname, targetPath, listName, view, accessToken, setLogs, customFormatter);
        setLogs((prev) => [...prev, `✅ Created view: ${view.Title}`]);
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `⚠️ Failed to create view ${view.Title}: ${errorMessage}`]);
      }
    }

    // Find the AssignedTo field
    const assignedToField = fields.find(f => 
      f.TypeAsString.toLowerCase().includes('user') || 
      f.Title.toLowerCase().includes('assigned')
    );

    // Build the select and expand query parameters
    const selectFields = ['Title', ...fields.map(f => f.InternalName)];
    let expandFields = '';
    
    if (assignedToField) {
      // For People Lookup fields, we need to select the Id and expand the field
      selectFields.push(`${assignedToField.InternalName}/Id`);
      expandFields = `&$expand=${assignedToField.InternalName}`;
    }

    // First, let's try to get the items without the AssignedTo field to see if that works
    const itemsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/items?$select=Id,${selectFields.join(',')}${expandFields}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!itemsRes.ok) {
      const errorText = await itemsRes.text();
      setLogs((prev) => [...prev, `⚠️ Failed to fetch items: ${errorText}`]);
      throw new Error(`Failed to fetch items: ${errorText}`);
    }

    const itemsData = await itemsRes.json() as SharePointResponse<ListItemBasic>;
    if (!itemsData.d?.results) {
      setLogs((prev) => [...prev, "No items found in source list"]);
      return;
    }

    const items = itemsData.d.results;
    setLogs((prev) => [...prev, `Found ${items.length} items to copy`]);

    // Create each item
    for (const item of items) {
      try {
        if (!item.Id) {
          setLogs((prev) => [...prev, `⚠️ Skipping item '${item.Title}' - no ID found`]);
          continue;
        }
        // Get the full item data including all custom fields
        const fullItemRes = await fetch(
          `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${item.Id})?$select=${selectFields.join(',')}${expandFields}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
            },
          }
        );
        if (!fullItemRes.ok) {
          const errorText = await fullItemRes.text();
          setLogs((prev) => [...prev, `⚠️ Failed to fetch full item data for item ${item.Id}: ${errorText}`]);
          continue;
        }
        const fullItem = await fullItemRes.json() as { d: ListItemFull };
        // Clean up the item data by removing SharePoint metadata and undefined values
        const cleanItem = Object.entries(fullItem.d)
          .filter(([key, value]) => {
            return !key.startsWith('_') && !key.startsWith('__') && value !== undefined && value !== null;
          })
          .reduce((acc, [key, value]) => {
            // Only include AssignedtoId, not Assignedto
            if (key === 'Assignedto') {
              // skip
            } else {
              acc[key] = value;
            }
            return acc;
          }, {} as Record<string, any>);
        const itemPayload = {
          __metadata: { type: "SP.ListItem" },
          ...cleanItem
        };
        const createItemRes = await fetch(
          `https://${targetHostname}${targetPath}/_api/web/lists/GetByTitle('${listName}')/items`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
            },
            body: JSON.stringify(itemPayload),
          }
        );
        if (createItemRes.ok) {
          setLogs((prev) => [...prev, `✅ Created item: ${item.Title || 'Untitled'}`]);
        } else {
          const errorText = await createItemRes.text();
          setLogs((prev) => [...prev, `⚠️ Failed to create item: ${errorText}`]);
        }
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `❌ Error creating item: ${errorMessage}`]);
      }
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Error cloning list '${listName}': ${errorMessage}`);
  }
}

export default function SharePointCloner() {
  const { instance, accounts } = useMsal();

  const [sourceUrl, setSourceUrl] = useState("");
  const [targetUrl, setTargetUrl] = useState("");
  const [cloneLists, setCloneLists] = useState(true);
  const [cloneLibraries, setCloneLibraries] = useState(true);
  const [clonePages, setClonePages] = useState(true);
  const [cloneSite, setCloneSite] = useState(false);
  const [logs, setLogs] = useState<string[]>(["Waiting for user..."]);
  const [items, setItems] = useState<SiteItem[]>([]);

  const toggleItem = (index: number) => {
    const updated = [...items];
    updated[index].selected = !updated[index].selected;
    setItems(updated);
  };

  const loadItemsFromSite = async () => {
    if (accounts.length === 0) {
      setLogs((prev) => [...prev, "User not signed in."]);
      return;
    }

    try {
      const response = await instance.acquireTokenSilent({
        scopes: ["Sites.Read.All"],
        account: accounts[0],
      });

      const accessToken = response.accessToken;
      const siteHostname = new URL(sourceUrl).hostname;
      const sitePath = new URL(sourceUrl).pathname;

      const siteResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );

      const siteData = await siteResp.json();
      const siteId = siteData.id;

      const [listsResp, pagesResp] = await Promise.all([
        fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists`, {
          headers: { Authorization: `Bearer ${accessToken}` },
        }),
        fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/pages`, {
          headers: { Authorization: `Bearer ${accessToken}` },
        }),
      ]);

      const listData = await listsResp.json();
      const pageData = await pagesResp.json();

      const listItems = listData.value.map((list: GraphListItem) => ({
        name: list.name,
        type: list.list?.template === "documentLibrary" ? "Document Library" : "List",
        selected: true,
      }));

      const pageItems = pageData.value.map((page: GraphPageItem) => ({
        name: page.name,
        type: "Page",
        selected: true,
      }));

      const combinedItems = [...listItems, ...pageItems];

      setItems(combinedItems);
      setLogs((prev) => [...prev, `Loaded ${combinedItems.length} items from source site.`]);
    } catch (error: unknown) {
      if (error instanceof InteractionRequiredAuthError) {
        instance.acquireTokenPopup({
          scopes: ["Sites.Read.All"],
        });
      } else {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `Error loading site items: ${errorMessage}`]);
      }
    }
  };

  const startCloning = async () => {
    setLogs((prev) => [...prev, "Starting clone..."]);

    if (accounts.length === 0) {
      setLogs((prev) => [...prev, "Cannot clone: user not signed in."]);
      return;
    }

    try {
      const sourceHostname = new URL(sourceUrl).hostname;
      const sourcePath = new URL(sourceUrl).pathname;
      const targetHostname = new URL(targetUrl).hostname;
      const targetPath = new URL(targetUrl).pathname;
      
      const response = await instance.acquireTokenSilent({
        scopes: [`https://${sourceHostname}/.default`],
        account: accounts[0],
      });

      const accessToken = response.accessToken;

      for (const item of items) {
        if (!item.selected) continue;

        if (item.type === "List" && cloneLists) {
          setLogs((prev) => [...prev, `Cloning list: ${item.name}`]);
          try {
            await cloneList(sourceHostname, sourcePath, targetHostname, targetPath, item.name, accessToken, setLogs);
          } catch (error: unknown) {
            const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
            setLogs((prev) => [...prev, `❌ ${errorMessage}`]);
          }
        } else if (item.type === "Document Library" && cloneLibraries) {
          setLogs((prev) => [...prev, `Cloning library: ${item.name}`]);
          // Placeholder: implement document library clone logic
        } else if (item.type === "Page" && clonePages) {
          setLogs((prev) => [...prev, `Cloning page: ${item.name}`]);
          // Placeholder: implement page cloning
        }
      }

      setLogs((prev) => [...prev, "✅ Cloning completed."]);
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      setLogs((prev) => [...prev, `❌ Cloning failed: ${errorMessage}`]);
    }
  };

  return (
    <div className="p-6 space-y-6 max-w-4xl mx-auto">
      <Card>
        <CardContent className="p-6 flex justify-center">
          <AuthButtons />
        </CardContent>
      </Card>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <Card>
          <CardContent className="p-6 space-y-2">
            <label className="font-semibold">🔍 Source Site URL</label>
            <Input value={sourceUrl} onChange={(e) => setSourceUrl(e.target.value)} />
            <Button onClick={loadItemsFromSite} className="mt-2">Load Site Items</Button>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="p-6 space-y-2">
            <label className="font-semibold">🎯 Target Site URL</label>
            <Input value={targetUrl} onChange={(e) => setTargetUrl(e.target.value)} />
          </CardContent>
        </Card>
      </div>

      <Card>
        <CardContent className="p-6 space-y-4">
          <label className="font-semibold">✅ What to Clone?</label>
          <div className="grid grid-cols-2 gap-4">
            <div className="flex items-center space-x-2">
              <Checkbox checked={cloneLists} onCheckedChange={val => typeof val === 'boolean' && setCloneLists(val)} />
              <span>Lists</span>
            </div>
            <div className="flex items-center space-x-2">
              <Checkbox checked={cloneLibraries} onCheckedChange={val => typeof val === 'boolean' && setCloneLibraries(val)} />
              <span>Document Libraries</span>
            </div>
            <div className="flex items-center space-x-2">
              <Checkbox checked={clonePages} onCheckedChange={val => typeof val === 'boolean' && setClonePages(val)} />
              <span>Pages</span>
            </div>
            <div className="flex items-center space-x-2">
              <Checkbox checked={cloneSite} onCheckedChange={val => typeof val === 'boolean' && setCloneSite(val)} />
              <span>Full Site</span>
            </div>
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardContent className="p-6 space-y-2">
          <label className="font-semibold">📂 Available Items</label>
          <div className="flex items-center space-x-4 mb-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => setItems(items.map(item => ({ ...item, selected: true })))}
            >
              Select All
            </Button>
            <Button
              variant="outline"
              size="sm"
              onClick={() => setItems(items.map(item => ({ ...item, selected: false })))}
            >
              Deselect All
            </Button>
          </div>
          {items.map((item, index) => (
            <div key={index} className="flex items-center space-x-2">
              <Checkbox
                checked={item.selected}
                onCheckedChange={() => toggleItem(index)}
              />
              <span>{item.name} ({item.type})</span>
            </div>
          ))}
        </CardContent>
      </Card>

      <Card>
        <CardContent className="p-6 space-y-4">
          <Button onClick={startCloning}>▶️ Start Cloning</Button>
          <div className="pt-2">
            <label className="font-semibold">🔧 Logs:</label>
            <div className="bg-gray-100 p-2 rounded text-sm h-40 overflow-y-scroll">
              {logs.map((log, i) => (
                <div key={i}>&gt; {log}</div>
              ))}
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}