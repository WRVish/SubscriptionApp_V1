import * as React from 'react';
import styles from './SubscriptionUpdates.module.scss';
import type { ISubscriptionUpdatesProps } from './ISubscriptionUpdatesProps';
import { Checkbox, MessageBar, MessageBarType, Spinner, SpinnerSize, Label } from '@fluentui/react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/site-users";

export interface IChoiceOption {
  key: string;
  text: string;
}

export interface ISubscriptionUpdatesState {
  availableChoices: IChoiceOption[];
  selectedChoices: string[];
  userItemId: number | null;
  loading: boolean;
  error: string | null;
  successMessage: string | null;
  isMultiChoice: boolean;
  referenceValue: string | null;
}

export default class SubscriptionUpdates extends React.Component<ISubscriptionUpdatesProps, ISubscriptionUpdatesState> {
  private _notificationTimeout: number | undefined;

  constructor(props: ISubscriptionUpdatesProps) {
    super(props);
    this.state = {
      availableChoices: [],
      selectedChoices: [],
      userItemId: null,
      loading: true,
      error: null,
      successMessage: null,
      isMultiChoice: false,
      referenceValue: null
    };
  }

  public componentDidMount(): void {
    this._loadData().catch(err => console.error(err));
  }

  public componentDidUpdate(prevProps: ISubscriptionUpdatesProps): void {
    if (prevProps.listId !== this.props.listId || prevProps.siteUrl !== this.props.siteUrl) {
      this._loadData().catch(err => console.error(err));
    }
  }

  public componentWillUnmount(): void {
    if (this._notificationTimeout) {
      globalThis.clearTimeout(this._notificationTimeout);
    }
  }

  private readonly _clearNotifications = (): void => {
    if (this._notificationTimeout) {
      globalThis.clearTimeout(this._notificationTimeout);
    }
    this.setState({ successMessage: null, error: null });
  }

  private readonly _showNotification = (type: 'success' | 'error', message: string): void => {
    this._clearNotifications();

    if (type === 'success') {
      this.setState({ successMessage: message });
    } else {
      this.setState({ error: message });
    }

    this._notificationTimeout = globalThis.setTimeout(() => {
      this.setState({ successMessage: null, error: null });
    }, 5000) as any;
  }

  private _replaceUserInfo(text: string): string {
    if (!text) return '';
    // Use replaceAll to satisfy SonarLint and modernize the replacement logic
    return text.replaceAll('<userinfo>', this.props.userEmail);
  }

  private async _loadData(): Promise<void> {
    if (!this.props.listId || !this.props.sp) {
      console.warn("Skipping data load: listId or sp instance missing.");
      this.setState({ loading: false });
      return;
    }

    this.setState({ loading: true, error: null, successMessage: null });

    try {
      this._logDiagnosticInfo();

      // 1. Fetch field info
      const field = await this.props.sp.web.lists.getById(this.props.listId).fields.getByInternalNameOrTitle(this.props.subscribeForColumn)();
      const choices = field.Choices || [];
      const isMultiChoice = field.TypeAsString === "MultiChoice";

      // 2. Build and execute query
      const { items, filterStr } = await this._fetchSubscriptionItems();

      console.log(`[SubscriptionUpdates] Executing Query: Filter="${filterStr}"`);
      console.log(`[SubscriptionUpdates] Items found: ${items.length}`);

      let userItemId = null;
      let selectedChoices: string[] = [];

      if (items.length > 0) {
        userItemId = items[0].Id;
        const value = items[0][this.props.subscribeForColumn];
        console.log(`[SubscriptionUpdates] Existing Item ID: ${userItemId}, Raw Value:`, value);
        selectedChoices = this._parseChoiceValue(value);
      } else {
        console.log("[SubscriptionUpdates] No existing subscription item found for this user.");
        if (this.props.referenceIdColumn) {
          console.warn(`[SubscriptionUpdates] You are searching by the Text column '${this.props.referenceIdColumn}'. If this column is empty in your list, please clear the 'Unique ID Column' setting in the web part properties to use the automatic Person-based matching (this.props.subscribedByColumn).`);
        }
      }

      this.setState({
        availableChoices: choices.map((c: string) => ({ key: c, text: c })),
        selectedChoices,
        userItemId,
        isMultiChoice,
        loading: false
      });
    } catch (error: any) {
      this._handleLoadError(error);
    }
  }

  private _logDiagnosticInfo(): void {
    console.log(`[SubscriptionUpdates] Loading data for List: ${this.props.listId} on Site: ${this.props.siteUrl || 'Current'}`);
    console.log(`[SubscriptionUpdates] User Identifier (Email): ${this.props.userEmail}`);
    console.log(`[SubscriptionUpdates] Column Mapping: SubscribeFor=${this.props.subscribeForColumn}, SubscribedBy=${this.props.subscribedByColumn}`);
  }

  private async _fetchSubscriptionItems(): Promise<{ items: any[], filterStr: string }> {
    const selectFields = ["Id", this.props.subscribeForColumn];
    let filterStr = "";
    let usePersonFilter = false;
    let targetColumnForId = this.props.subscribedByColumn;

    // 1. Determine if we are using the Unique ID (Text) or a Person column
    if (this.props.referenceIdColumn && this.props.referenceIdColumn !== this.props.subscribedByColumn) {
      try {
        // Fetch field metadata to see if it's a Person field
        const refField = await this.props.sp.web.lists.getById(this.props.listId).fields.getByInternalNameOrTitle(this.props.referenceIdColumn)();

        if (refField.TypeAsString === "User") {
          usePersonFilter = true;
          targetColumnForId = this.props.referenceIdColumn;
          console.log(`[SubscriptionUpdates] Unique ID Column '${this.props.referenceIdColumn}' is a Person field. Using ID-based filtering.`);
        } else {
          selectFields.push(this.props.referenceIdColumn);
          filterStr = `${this.props.referenceIdColumn} eq '${this.props.userEmail}'`;
          console.log(`[SubscriptionUpdates] Unique ID Column '${this.props.referenceIdColumn}' is a Text-like field. Using direct string comparison.`);
        }
      } catch (err) {
        // Fallback to text matching if metadata fetch fails
        console.warn(`[SubscriptionUpdates] Failed to fetch field metadata for '${this.props.referenceIdColumn}', falling back to text comparison:`, err);
        selectFields.push(this.props.referenceIdColumn);
        filterStr = `${this.props.referenceIdColumn} eq '${this.props.userEmail}'`;
      }
    } else {
      // Default path: use the main Person column
      usePersonFilter = true;
      targetColumnForId = this.props.subscribedByColumn;
    }

    // 2. If filtering by a Person column, use the internal User ID for robustness
    if (usePersonFilter) {
      const userResult = await this.props.sp.web.ensureUser(this.props.userEmail);
      const userId = userResult.Id;
      filterStr = `${targetColumnForId}Id eq ${userId}`;
      console.log(`[SubscriptionUpdates] Executing ID-based filter: ${filterStr}`);
    }

    const items = await this.props.sp.web.lists.getById(this.props.listId).items
      .select(...selectFields)
      .filter(filterStr)();

    return { items, filterStr };
  }

  private _parseChoiceValue(value: any): string[] {
    if (Array.isArray(value)) return value;
    if (typeof value === 'string') return [value];
    return [];
  }

  private _handleLoadError(error: any): void {
    console.error("[SubscriptionUpdates] Error loading subscription data", error);
    let errorMsg = "Could not load subscription data. Please check configuration.";
    if (error?.data?.responseBody?.error) {
      errorMsg += ` ${error.data.responseBody.error.message.value}`;
    }
    this.setState({ error: errorMsg, loading: false });
  }

  private readonly _onCheckboxChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean, choiceKey?: string): void => {
    if (choiceKey !== undefined) {
      this.setState((prevState) => {
        let selectedChoices = [...prevState.selectedChoices];
        if (isChecked) {
          if (prevState.isMultiChoice) {
            if (!selectedChoices.includes(choiceKey)) {
              selectedChoices.push(choiceKey);
            }
          } else {
            // If it's single choice, only one can be selected
            selectedChoices = [choiceKey];
          }
        } else {
          selectedChoices = selectedChoices.filter(c => c !== choiceKey);
        }
        return { selectedChoices };
      });
    }
  }

  private readonly _onUpdateClick = async (): Promise<void> => {
    if (!this.props.listId || !this.props.sp) return;

    this.setState({ loading: true, error: null, successMessage: null });

    try {
      const updateData = this._prepareUpdatePayload();

      if (this.state.userItemId) {
        await this._executeUpdate(updateData);
      } else {
        await this._executeAdd(updateData);
      }

      await this._loadData();
      this.setState({ loading: false });

      if (this.props.showSuccessNotification) {
        this._showSuccessMessage();
      }
    } catch (error: any) {
      this._handleUpdateError(error);
    }
  }

  private _prepareUpdatePayload(): any {
    const updateData: any = {};

    // Prevent explicitly setting system date fields which throws errors for standard users
    if (this.props.subscribedOnColumn && this.props.subscribedOnColumn !== "Created" && this.props.subscribedOnColumn !== "Modified") {
      updateData[this.props.subscribedOnColumn] = new Date().toISOString();
    }

    if (this.state.isMultiChoice) {
      updateData[this.props.subscribeForColumn] = this.state.selectedChoices;
    } else {
      updateData[this.props.subscribeForColumn] = this.state.selectedChoices.length > 0 ? this.state.selectedChoices[0] : null;
    }

    return updateData;
  }

  private async _executeUpdate(updateData: any): Promise<void> {
    console.log(`[SubscriptionUpdates] Updating item ID ${this.state.userItemId} with Data:`, updateData);
    await this.props.sp.web.lists.getById(this.props.listId).items.getById(this.state.userItemId!).update(updateData);
  }

  private async _executeAdd(updateData: any): Promise<void> {
    const user = await this.props.sp.web.ensureUser(this.props.userEmail);
    
    // Explicitly setting AuthorId throws an Access Denied error for users with item-level permissions.
    // SharePoint automatically sets Author (Created By) when a user creates an item.
    if (this.props.subscribedByColumn && this.props.subscribedByColumn !== "Author") {
      updateData[`${this.props.subscribedByColumn}Id`] = user.Id;
    }

    if (this.props.referenceIdColumn) {
      if (this.props.referenceIdColumn === this.props.subscribedByColumn) {
        // Already handled above, no need to set it again as a string
      } else {
        try {
          const refField = await this.props.sp.web.lists.getById(this.props.listId).fields.getByInternalNameOrTitle(this.props.referenceIdColumn)();
          if (refField.TypeAsString === "User") {
            updateData[`${this.props.referenceIdColumn}Id`] = user.Id;
          } else {
            updateData[this.props.referenceIdColumn] = this.props.userEmail;
          }
        } catch (e) {
          // Fallback if we cannot get field metadata
          updateData[this.props.referenceIdColumn] = this.props.userEmail;
        }
      }
    }

    // Always provide a Title for new items as it's a required field by default in SharePoint lists.
    updateData['Title'] = `Subscription for ${this.props.userDisplayName || this.props.userEmail}`;

    console.log(`[SubscriptionUpdates] Adding new item with Data:`, updateData);
    await this.props.sp.web.lists.getById(this.props.listId).items.add(updateData);
  }

  private _showSuccessMessage(): void {
    const successMsg = this.props.successNotification ? this._replaceUserInfo(this.props.successNotification) : "Subscription updated successfully!";
    this._showNotification('success', successMsg);
  }

  private _handleUpdateError(error: any): void {
    console.error("[SubscriptionUpdates] Error updating subscription", error);
    let errorMsg = "Failed to update subscription. Please try again.";
    if (error?.data?.responseBody?.error) {
      errorMsg += ` Detail: ${error.data.responseBody.error.message.value}`;
    }
    this._showNotification('error', errorMsg);
    this.setState({ loading: false });
  }

  public render(): React.ReactElement<ISubscriptionUpdatesProps> {
    const { listId, customTitle, showTitle, descriptionText, buttonText, titleFontSize, descriptionFontSize, notificationFontSize, buttonFontSize, choicesFontSize, choicesLabel, titleBarOpacity } = this.props;
    const { availableChoices, selectedChoices, loading, error, successMessage } = this.state;

    if (!listId) {
      return (
        <section className={styles.subscriptionUpdates}>
          <div className={styles.container}>
            <MessageBar messageBarType={MessageBarType.info}>
              Please configure the web part properties.
            </MessageBar>
          </div>
        </section>
      );
    }

    return (
      <section className={styles.subscriptionUpdates}>
        <div className={styles.container}>
          {showTitle && (
            <div className={styles.titleBar} style={{ opacity: titleBarOpacity / 100 }}>
              <h2 className={styles.title} style={{ fontSize: `${titleFontSize}px` }}>
                {customTitle || "Subscription Updates"}
              </h2>
            </div>
          )}

          <div className={styles.content}>
            {loading && <Spinner size={SpinnerSize.large} label="Processing..." />}

            {(error || successMessage) && (
              <div className={styles.notifications} style={{ width: '100%', marginBottom: '10px' }}>
                {error && (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={this._clearNotifications}
                    dismissButtonAriaLabel="Close"
                  >
                    <span style={{ fontSize: `${notificationFontSize}px` }}>{error}</span>
                  </MessageBar>
                )}
                {successMessage && (
                  <MessageBar
                    messageBarType={MessageBarType.success}
                    isMultiline={false}
                    onDismiss={this._clearNotifications}
                    dismissButtonAriaLabel="Close"
                  >
                    <span style={{ fontSize: `${notificationFontSize}px` }}>{successMessage}</span>
                  </MessageBar>
                )}
              </div>
            )}

            {descriptionText && (
              <div className={styles.header} style={{ marginBottom: '15px', width: '100%' }}>
                <p className={styles.description} style={{ fontSize: `${descriptionFontSize}px`, margin: 0, padding: 0 }}>
                  {this._replaceUserInfo(descriptionText)}
                </p>
                <div style={{ fontSize: '12px', color: '#666', marginTop: '5px' }}>
                  Viewing subscriptions for: <strong>{this.props.userEmail}</strong>
                </div>
              </div>
            )}

            {!loading && (
              <>
                <div className={styles.details} style={{ width: '100%', maxWidth: '400px', textAlign: 'left', marginBottom: '20px' }}>
                  <Label style={{ fontSize: `${choicesFontSize}px` }}>{choicesLabel || "Subscribe For"}</Label>
                  {availableChoices.map(option => (
                    <Checkbox
                      key={option.key}
                      label={option.text}
                      checked={selectedChoices.includes(option.key)}
                      onChange={(ev, checked) => this._onCheckboxChange(ev, checked, option.key)}
                      styles={{ root: { marginBottom: '10px' } }}
                    />
                  ))}
                </div>
                <button
                  className={styles.actionButton}
                  onClick={this._onUpdateClick}
                  disabled={loading}
                  style={{ fontSize: `${buttonFontSize}px` }}
                >
                  {buttonText || "Update"}
                </button>
              </>
            )}
          </div>
        </div>
      </section>
    );
  }
}
