/**
 * TypeScript types for SharePoint column schema
 * Used by the Column Inspector feature
 */

export enum FieldType {
  Text = 'Text',
  Note = 'Note',
  Number = 'Number',
  Currency = 'Currency',
  DateTime = 'DateTime',
  Choice = 'Choice',
  Lookup = 'Lookup',
  Boolean = 'Boolean',
  User = 'User',
  Calculated = 'Calculated',
  Computed = 'Computed',
  URL = 'URL',
  Integer = 'Integer',
  Counter = 'Counter',
  Guid = 'Guid',
  MultiChoice = 'MultiChoice',
  TaxonomyFieldType = 'TaxonomyFieldType',
  Location = 'Location',
  Geolocation = 'Geolocation',
  Image = 'Image',
  Thumbnail = 'Thumbnail',
}

export interface ColumnSchema {
  /** Internal name of the column */
  internalName: string;

  /** Display name of the column */
  title: string;

  /** Column type */
  type: string;

  /** Column GUID */
  id: string;

  /** Whether the column is required */
  required: boolean;

  /** Whether the column is hidden */
  hidden: boolean;

  /** Whether the column is read-only */
  readOnly: boolean;

  /** Column description */
  description?: string;

  /** Default value */
  defaultValue?: string;

  /** For calculated columns: formula */
  formula?: string;

  /** For lookup columns: source list ID */
  lookupListId?: string;

  /** For lookup columns: source field name */
  lookupField?: string;

  /** For choice/multi-choice: available choices */
  choices?: string[];

  /** For taxonomy columns: term set ID */
  termSetId?: string;

  /** Maximum length for text fields */
  maxLength?: number;

  /** Whether to index this column */
  indexed: boolean;

  /** Static name (used in views/CAML queries) */
  staticName?: string;

  /** Group/category this field belongs to */
  group?: string;
}

export interface ListSchema {
  /** List/Library title */
  title: string;

  /** List internal name */
  internalName: string;

  /** List GUID */
  id: string;

  /** List template type */
  baseTemplate: number;

  /** List template name */
  baseType: string;

  /** List description */
  description?: string;

  /** All columns in the list */
  fields: ColumnSchema[];

  /** Content types associated with this list */
  contentTypes?: ContentTypeInfo[];

  /** Whether content types are enabled/managed in this list */
  contentTypesEnabled?: boolean;

  /** Whether versioning is enabled */
  enableVersioning: boolean;

  /** Whether the list is hidden */
  hidden: boolean;

  /** Server relative URL */
  serverRelativeUrl: string;

  /** Web absolute URL (for generating management links) */
  webAbsoluteUrl?: string;
}

export interface ContentTypeInfo {
  /** Content type ID */
  id: string;

  /** Content type name */
  name: string;

  /** Content type description */
  description?: string;

  /** Parent content type ID */
  parent?: string;

  /** Fields in this content type */
  fields?: string[];
}

export interface ColumnInspectorState {
  /** Current list being inspected */
  currentList: ListSchema | null;

  /** Whether data is being fetched */
  loading: boolean;

  /** Error message if fetch failed */
  error: string | null;

  /** Filter text for column search */
  filterText: string;

  /** Whether to show hidden columns */
  showHidden: boolean;

  /** Whether inspector panel is visible */
  visible: boolean;
}

export interface SharePointFieldData {
  /** Raw field data from SharePoint REST API */
  EntityPropertyName: string;
  FieldTypeKind: number;
  Hidden: boolean;
  Id: string;
  Indexed: boolean;
  InternalName: string;
  ReadOnlyField: boolean;
  Required: boolean;
  StaticName: string;
  Title: string;
  TypeAsString: string;
  DefaultValue?: string;
  Description?: string;
  Group?: string;
  MaxLength?: number;
  Choices?: { results: string[] };
  Formula?: string;
  LookupList?: string;
  LookupField?: string;
  TermSetId?: string;
}
