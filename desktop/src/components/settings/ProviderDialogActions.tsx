import { ProviderTestButton } from "./ProviderDialogFields";

interface AddProviderDialogActionsProps {
  canAdd: boolean;
  testing: boolean;
  cancelLabel: string;
  addLabel: string;
  addAndTestLabel: string;
  testingLabel: string;
  onCancel: () => void;
  onAdd: () => void;
  onAddWithTest: () => void;
}

export function AddProviderDialogActions({
  canAdd,
  testing,
  cancelLabel,
  addLabel,
  addAndTestLabel,
  testingLabel,
  onCancel,
  onAdd,
  onAddWithTest,
}: AddProviderDialogActionsProps) {
  return (
    <>
      <button className="btn-secondary" onClick={onCancel}>
        {cancelLabel}
      </button>
      <button className="btn-primary" onClick={onAdd} disabled={!canAdd}>
        {addLabel}
      </button>
      <ProviderTestButton
        className="btn-primary"
        testing={testing}
        label={addAndTestLabel}
        testingLabel={testingLabel}
        disabled={!canAdd || testing}
        onClick={onAddWithTest}
      />
    </>
  );
}

interface EditProviderDialogActionsProps {
  testing: boolean;
  canTest: boolean;
  testConnectionLabel: string;
  cancelLabel: string;
  saveLabel: string;
  onTest: () => void;
  onCancel: () => void;
  onSave: () => void;
}

export function EditProviderDialogActions({
  testing,
  canTest,
  testConnectionLabel,
  cancelLabel,
  saveLabel,
  onTest,
  onCancel,
  onSave,
}: EditProviderDialogActionsProps) {
  return (
    <>
      <ProviderTestButton
        className="btn-test-full"
        testing={testing}
        label={testConnectionLabel}
        disabled={!canTest || testing}
        onClick={onTest}
      />
      <div className="dialog-actions-right">
        <button className="btn-secondary" onClick={onCancel}>
          {cancelLabel}
        </button>
        <button className="btn-primary" onClick={onSave}>
          {saveLabel}
        </button>
      </div>
    </>
  );
}
