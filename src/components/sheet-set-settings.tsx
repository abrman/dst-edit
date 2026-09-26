import { useState } from "react";
import type React from "react";
import { Plus, Trash2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { CommitInput } from "@/components/commit-input";
import type { CustomPropDef, DstModel } from "@/lib/dst-model";
import { mapSheets } from "@/lib/tree";
import { environPath, relativePath } from "@/lib/paths";

interface SheetSetSettingsProps {
  model: DstModel;
  baseDir: string | null;
  onChange: (update: (model: DstModel) => DstModel) => void;
}

// The sheet set properties AutoCAD shows in Sheet Set Properties, in its order
const SET_FIELDS: { name: string; label: string; hint?: string }[] = [
  { name: "Name", label: "Sheet set name", hint: "The title AutoCAD's Sheet Set Manager shows for this sheet set." },
  { name: "Desc", label: "Description" },
  { name: "ProjectNumber", label: "Project number" },
  { name: "ProjectName", label: "Project name" },
  { name: "ProjectPhase", label: "Project phase" },
  { name: "ProjectMilestone", label: "Project milestone" },
];

const fieldClass = "h-9 rounded-md border-input bg-background px-3 shadow-xs hover:border-input";

function Section({ title, description, children }: { title: string; description?: string; children: React.ReactNode }) {
  return (
    <section className="grid gap-4 rounded-lg border p-5 md:grid-cols-[16rem_1fr] md:gap-8">
      <div>
        <h3 className="font-semibold">{title}</h3>
        {description && <p className="mt-1 text-sm text-muted-foreground">{description}</p>}
      </div>
      <div className="grid gap-4">{children}</div>
    </section>
  );
}

function Field({ label, hint, children }: { label: string; hint?: React.ReactNode; children: React.ReactNode }) {
  return (
    <div className="grid gap-1.5">
      <Label>{label}</Label>
      {children}
      {hint && <p className="text-xs text-muted-foreground">{hint}</p>}
    </div>
  );
}

export function SheetSetSettings({ model, baseDir, onChange }: SheetSetSettingsProps) {
  const [newName, setNewName] = useState("");
  const [newScope, setNewScope] = useState<CustomPropDef["scope"]>("sheet");
  const [newValue, setNewValue] = useState("");
  const [removing, setRemoving] = useState<CustomPropDef | null>(null);
  const [showPaths, setShowPaths] = useState(false);

  const setProp = (name: string, value: string) =>
    onChange((m) => ({ ...m, setProps: { ...m.setProps, [name]: value } }));

  const otherFields = Object.keys(model.setProps).filter((name) => !SET_FIELDS.some((f) => f.name === name));

  const setOutputFolder = (fileName: string) =>
    onChange((m) => ({
      ...m,
      outputDir: {
        fileName,
        relative: fileName && baseDir ? relativePath(baseDir, fileName) : m.outputDir.relative,
        environ: fileName ? environPath(fileName) : "",
      },
    }));

  const takenNames = new Set([...model.customProps.map((d) => d.name), ...model.sheetColumns]);
  const nameError =
    newName.trim() && takenNames.has(newName.trim()) ? "A property with this name already exists." : null;

  const addProperty = () => {
    const name = newName.trim();
    if (!name || nameError) return;
    onChange((m) => ({
      ...m,
      customProps: [...m.customProps, { name, scope: newScope, value: newValue }],
      // Like AutoCAD, a new sheet property is added to every sheet with the default value
      children:
        newScope === "sheet"
          ? mapSheets(m.children, (s) => ({ ...s, props: { ...s.props, [name]: newValue } }))
          : m.children,
    }));
    setNewName("");
    setNewValue("");
  };

  const renameProperty = (def: CustomPropDef, name: string) => {
    name = name.trim();
    if (!name || name === def.name || takenNames.has(name)) return;
    onChange((m) => ({
      ...m,
      customProps: m.customProps.map((d) => (d.name === def.name ? { ...d, name } : d)),
      children: mapSheets(m.children, (s) => {
        if (!(def.name in s.props)) return s;
        const { [def.name]: value, ...rest } = s.props;
        return { ...s, props: { ...rest, [name]: value } };
      }),
    }));
  };

  const removeProperty = (def: CustomPropDef) =>
    onChange((m) => ({
      ...m,
      customProps: m.customProps.filter((d) => d.name !== def.name),
      children: mapSheets(m.children, (s) => {
        if (!(def.name in s.props)) return s;
        const { [def.name]: _removed, ...rest } = s.props;
        return { ...s, props: rest };
      }),
    }));

  return (
    <div className="space-y-4">
      <Section title="Sheet set" description="General information stored in the sheet set.">
        {[...SET_FIELDS, ...otherFields.map((name) => ({ name, label: name, hint: undefined }))].map((field) => (
          <Field key={field.name} label={field.label} hint={field.hint}>
            <CommitInput
              value={model.setProps[field.name] ?? ""}
              onCommit={(v) => setProp(field.name, v)}
              className={fieldClass}
            />
          </Field>
        ))}
      </Section>

      <Section title="Publishing" description="Where AutoCAD saves PDFs and plots when you publish the sheet set.">
        <Field
          label="Default output folder"
          hint={
            baseDir ? (
              <>
                The sheet set's own folder is <code className="break-all">{baseDir}</code> (worked out from the drawing
                paths), so the relative path is filled in for you.
              </>
            ) : (
              "The sheet set's own folder couldn't be worked out from the drawing paths, so check the relative path below."
            )
          }
        >
          <CommitInput
            value={model.outputDir.fileName}
            onCommit={setOutputFolder}
            placeholder="C:\Projects\Job\Plots"
            className={fieldClass}
          />
        </Field>

        <div className="rounded-md bg-muted/50 px-3 py-2 text-xs">
          <div className="flex items-center justify-between gap-2">
            <span className="text-muted-foreground">
              Also stored as{" "}
              <code className="break-all text-foreground">{model.outputDir.relative || "(no relative path)"}</code>{" "}
              and <code className="break-all text-foreground">{model.outputDir.environ || "(no environment path)"}</code>
            </span>
            <Button variant="link" size="sm" className="h-auto p-0 text-xs" onClick={() => setShowPaths((v) => !v)}>
              {showPaths ? "Hide" : "Edit"}
            </Button>
          </div>
          {showPaths && (
            <div className="mt-3 grid gap-3">
              <Field label="Relative to the sheet set" hint="AutoCAD uses this when the project folder moves.">
                <CommitInput
                  value={model.outputDir.relative}
                  onCommit={(relative) => onChange((m) => ({ ...m, outputDir: { ...m.outputDir, relative } }))}
                  className={fieldClass}
                />
              </Field>
              <Field label="With environment variables">
                <CommitInput
                  value={model.outputDir.environ}
                  onCommit={(environ) => onChange((m) => ({ ...m, outputDir: { ...m.outputDir, environ } }))}
                  className={fieldClass}
                />
              </Field>
            </div>
          )}
        </div>

        <Field label="Default PDF file name" hint="Used when publishing all sheets to a single file.">
          <CommitInput
            value={model.defaultFilename}
            onCommit={(defaultFilename) => onChange((m) => ({ ...m, defaultFilename }))}
            className={fieldClass}
          />
        </Field>
      </Section>

      <Section
        title="Custom properties"
        description="Sheet properties get a value on every sheet and show as columns on the Sheets tab. Sheet set properties have one value for the whole set."
      >
        <div className="overflow-hidden rounded-md border">
          <table className="w-full text-sm">
            <thead className="bg-muted/50 text-left text-xs text-muted-foreground">
              <tr>
                <th className="px-3 py-2 font-medium">Name</th>
                <th className="px-3 py-2 font-medium">Owner</th>
                <th className="px-3 py-2 font-medium">Value</th>
                <th className="w-10" />
              </tr>
            </thead>
            <tbody>
              {model.customProps.length === 0 && (
                <tr>
                  <td colSpan={4} className="px-3 py-4 text-center text-muted-foreground">
                    No custom properties yet.
                  </td>
                </tr>
              )}
              {model.customProps.map((def) => (
                <tr key={def.name} className="border-t">
                  <td className="px-1 py-1">
                    <CommitInput value={def.name} onCommit={(name) => renameProperty(def, name)} className="font-medium" />
                  </td>
                  <td className="px-3 py-1 whitespace-nowrap text-muted-foreground">
                    {def.scope === "sheet" ? "Sheet" : "Sheet set"}
                  </td>
                  <td className="px-1 py-1">
                    <CommitInput
                      value={def.value}
                      onCommit={(value) =>
                        onChange((m) => ({
                          ...m,
                          customProps: m.customProps.map((d) => (d.name === def.name ? { ...d, value } : d)),
                        }))
                      }
                      placeholder={def.scope === "sheet" ? "Default for new sheets" : "Value"}
                    />
                  </td>
                  <td className="pr-1">
                    <Button
                      variant="ghost"
                      size="icon-sm"
                      onClick={() => setRemoving(def)}
                      aria-label={`Remove ${def.name}`}
                      className="text-muted-foreground hover:text-destructive"
                    >
                      <Trash2 />
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <form
          className="grid gap-2 sm:grid-cols-[1fr_9rem_1fr_auto] sm:items-end"
          onSubmit={(e) => {
            e.preventDefault();
            addProperty();
          }}
        >
          <Field label="New property">
            <Input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="e.g. DRAWN_BY" aria-invalid={!!nameError} />
          </Field>
          <Field label="Owner">
            <Select value={newScope} onValueChange={(v) => setNewScope(v as CustomPropDef["scope"])}>
              <SelectTrigger className="w-full">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="sheet">Sheet</SelectItem>
                <SelectItem value="set">Sheet set</SelectItem>
              </SelectContent>
            </Select>
          </Field>
          <Field label={newScope === "sheet" ? "Default value" : "Value"}>
            <Input value={newValue} onChange={(e) => setNewValue(e.target.value)} />
          </Field>
          <Button type="submit" disabled={!newName.trim() || !!nameError}>
            <Plus /> Add
          </Button>
          {nameError && <p className="text-xs text-destructive sm:col-span-4">{nameError}</p>}
        </form>
      </Section>

      <AlertDialog open={removing !== null} onOpenChange={(open) => !open && setRemoving(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Remove the “{removing?.name}” property?</AlertDialogTitle>
            <AlertDialogDescription>
              {removing?.scope === "sheet"
                ? "It's removed from the sheet set and from every sheet, along with each sheet's value. Title blocks that show this property will go blank."
                : "It's removed from the sheet set. Title blocks that show this property will go blank."}
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>
            <AlertDialogAction
              className="bg-destructive text-white hover:bg-destructive/90"
              onClick={() => {
                if (removing) removeProperty(removing);
                setRemoving(null);
              }}
            >
              Remove
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
