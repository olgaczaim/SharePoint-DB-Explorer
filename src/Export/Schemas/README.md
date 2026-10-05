# Deployment XML schemas

These schemas are derived from Microsoft's published Open Specifications, retrieved on 2026-10-03. The independent exporter validates the package against them, then checks object relationships, required payload references, recovered lengths and SHA256. This does not prove that a target SharePoint farm can import the package.

The deployment schemas are [MS-PRIMEPF] Appendix A, sections 5.1–5.8:

- Manifest: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/80fcfb67-51ff-48cd-bcd0-2895c9c4cfe6
- ExportSettings: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/4a84999c-9e37-48f2-a592-74080bb2f317
- LookupListMap: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/057b696b-40f7-4ed0-bc87-25e23e551e75
- Requirements: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/aa4873c1-7ed4-4e73-824c-5da8e15c14fc
- RootObjectMap: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/b7a07fcc-9eef-42ff-b2c7-3b6caa2a2c04
- SystemData: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/7d695b74-5f35-43fb-b3c2-c141c9b3e917
- UserGroup: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/3cfdc65b-f65f-4313-9f15-6c2679bce4ba
- ViewFormsList: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/98b069c9-4c1d-4cd8-9c90-9e1362fc3f7f

Each wsswire include is the complete [MS-WSSCAML] Appendix A schema with its SharePoint SOAP namespace replaced by the containing deployment namespace, as instructed by [MS-PRIMEPF]. Source: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wsscaml/bc8fc7a6-9ba5-42f2-bdec-6261f6a00621

Microsoft permits redistribution of included schemas and code samples for implementations in its Open Specifications intellectual-property notice. These are not unmodified publisher artifacts. The following narrow corrections make the published schemas compile or cover documented current deployment object types; they do not add wildcard acceptance:

- Shared CAML: move invalid mixed="true" from the four Script/XslLink/JS/JSLink element declarations to their anonymous complexType. The element content models remain unchanged.
- Manifest SPForm.Type: retain exactly DisplayForm/EditForm/NewForm/empty, but restrict xs:string because the published core:FormType base excludes the expressly listed empty value.
- Manifest SPAlert: replace five undefined Guid references with core:UniqueIdentifierWithoutBraces, the deployment GUID type.
- RootObjectMap SPDeploymentObjectType: add List and Folder; ExportSettings SPExportObject.Type: add List, Folder and File. Microsoft's current SDK identifies these deployment object types: https://learn.microsoft.com/en-us/sharepoint/dev/schema/spdeploymentobjecttype-simple-type-deploymentrootobjectmap . A current Microsoft manifest example also includes a List root mapping: https://learn.microsoft.com/en-us/sharepoint/dev/apis/migration-manifest .

The writer exports current item/file/attachment content, typed supported fields, content types and referenced users. IncludeSecurity=None deliberately excludes ACLs and role assignments; workflow state is excluded. The current-only command uses IncludeVersions=CurrentVersion; the history command uses IncludeVersions=All and exports retained document and item versions. Historical attachment sets are excluded. Built-in storage-only FieldRef entries are recreated by the selected source list template; complete custom Field definitions are preserved. Missing required bytes or unsupported custom metadata cause the package to fail before publication.

The exporter selects a content-deployment compatibility profile from the detected source generation, separately from the source SQL build:

| Source generation | DatabaseVersion | Deployment schema / SiteVersion |
| --- | --- | --- |
| SharePoint Server 2016 | 38455 | 15.0.0.0 / 15 |
| SharePoint Server 2019 | 12710 | 15.0.0.0 / 15 |
| SharePoint Server Subscription Edition | 7123 | 15.0.0.0 / 15 |

The actual source build is retained in the audit. Unsupported generations are rejected. Current and history packages have been exercised against a restored Subscription Edition database; the 2016 and 2019 profiles have fixture coverage but have not been exercised against real databases of those generations. No target-farm import has been verified.

Uncompressed packages are intended for Import-SPWeb -NoFileCompression on a compatible target: https://learn.microsoft.com/en-us/powershell/module/sharepointserver/import-spweb?view=sharepoint-server-ps . Target templates, feature availability, source/target compatibility and actual import behavior remain to be verified. Source SQL paths may have no hostname; ExportSettings retains an actual absolute source URL when one is available and otherwise records the stored relative source path without inventing a farm host.

For templates 100/101 storage-only built-in field references without a recorded field GUID, the writer uses the explicitly published standard field identity facts from Microsoft's MS-PRIMEPF 3.1 Manifest example: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/b767a160-15b1-4842-8283-19dd1b0f4eb3 . This bounded identity profile covers Title, moderation comments, modified/created user display fields, file type, HTML file type, source/shared indexes, template URL and XML document metadata. It applies only to the exact generic-list/document-library template feature IDs. Source custom field identities are retained; an unknown non-null reference causes failure. Duplicate structural fields such as ContentTypeId are checked against the canonical ListItem attributes and represented there, without inventing an otherwise missing field GUID.

Source-to-deployment CAML conversion is deliberately narrow. The restored SE RTM source records `ContentType@DelayActivateTemplateBinding="GROUP,SPSPERS,SITEPAGEPUBLISHING"`, an attribute absent from the published [MS-WSSCAML] `ContentTypeDefinitionTP`. The writer omits only that exact value on standard Item/Document content types under feature `695b6570-a48b-4a8e-8ea5-26ea7fc1d162`, exact templates 100/101 feature IDs and the known RTM source profile. The original complete content-type and field XML and each conversion are retained in `package-audit.json`. Unknown marker values or content-type identities fail before publication; the XSD is unchanged. The conversion does not establish target-farm support for deferred template activation, and an actual import round trip remains unverified. Canonical content-type definitions: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wsscaml/bc8fc7a6-9ba5-42f2-bdec-6261f6a00621 .

Additional source settings are converted independently from the public MS-WSSFO3 List Flags definitions into the explicit canonical SPList settings: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wssfo3/eb951bd3-6653-4eb7-b708-77a41c0627ec and https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-primepf/98ec7c7e-21ea-44da-b528-96bf6c69394e . Versioning, checkout, drafts, attachments, content-type controls and folder creation therefore use the source configuration rather than target defaults. Raw Flags/Flags2 are also retained in the audit; Flags2 is not invented as an extension to the published XSD. The root-web-only list bit is represented in SystemData.RootWebOnlyLists using the list GUID. Sources marked as excluded from migration are rejected. Enabled moderation or list validation requiring approval state/formulas not available in this catalog also cause an explicit package capability error; normal file and attachment saving remains available.

SystemData declares exact stored source root/selected-web, parent-folder and UserInfo-list identities read from active SQL rows. Complete source context is required before export, retained in the audit and checked against package declarations. Missing context causes export to fail; no incomplete context is invented. The canonical SystemData schema is unchanged. Attachment XML also preserves recorded creation time and owner author/editor attribution. These corrections retain schema and semantic checks without asserting a verified target-farm import.
