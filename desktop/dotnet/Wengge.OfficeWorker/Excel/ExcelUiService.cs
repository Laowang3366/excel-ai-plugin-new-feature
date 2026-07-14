using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Excel;

internal sealed class ExcelUiService(ExcelSessionService sessions)
{
    private static readonly IReadOnlyDictionary<string, string> ControlTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["dropdown"] = "Forms.ComboBox.1",
        ["checkbox"] = "Forms.CheckBox.1",
        ["textbox"] = "Forms.TextBox.1",
        ["label"] = "Forms.Label.1",
        ["listbox"] = "Forms.ListBox.1",
        ["spinner"] = "Forms.SpinButton.1",
        ["scrollbar"] = "Forms.ScrollBar.1",
        ["optionbutton"] = "Forms.OptionButton.1",
        ["groupbox"] = "Forms.Frame.1",
    };

    public object AddControl(JsonElement parameters)
    {
        var sheetName = parameters.RequiredString("sheetName");
        var controlType = parameters.RequiredString("controlType");
        var name = parameters.RequiredString("name");
        var caption = parameters.OptionalString("caption") ?? name;
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheet = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            sheet = ((dynamic)workbook).Sheets.Item(sheetName);
            dynamic sheetApi = sheet;
            return controlType.Equals("button", StringComparison.OrdinalIgnoreCase)
                ? AddButton(sheetApi, parameters, name, caption)
                : AddActiveX(sheetApi, parameters, controlType, name, caption);
        }
        finally
        {
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    public object RemoveControl(string sheetName, string name)
    {
        return WithSheet(sheet =>
        {
            object? control = null;
            try
            {
                try { control = sheet.Buttons().Item(name); }
                catch { control = sheet.OLEObjects.Item(name); }
                if (control is null) throw new OfficeWorkerException("control_not_found", $"找不到控件: {name}");
                ((dynamic)control).Delete();
                return new { removed = true, name };
            }
            finally
            {
                ComInterop.Release(control);
            }
        }, sheetName);
    }

    public object ListControls(string sheetName)
    {
        return WithSheet(sheet =>
        {
            var controls = new List<object>();
            object? buttons = null;
            object? oleObjects = null;
            try
            {
                buttons = sheet.Buttons();
                dynamic buttonsApi = buttons;
                for (var index = 1; index <= Convert.ToInt32(buttonsApi.Count); index++)
                {
                    object? button = null;
                    try
                    {
                        button = buttonsApi.Item(index);
                        dynamic buttonApi = button;
                        controls.Add(new
                        {
                            name = Convert.ToString(buttonApi.Name),
                            controlType = "button",
                            controlKind = "form",
                            left = Convert.ToDouble(buttonApi.Left),
                            top = Convert.ToDouble(buttonApi.Top),
                            width = Convert.ToDouble(buttonApi.Width),
                            height = Convert.ToDouble(buttonApi.Height),
                            caption = Convert.ToString(buttonApi.Caption),
                            onAction = Convert.ToString(buttonApi.OnAction),
                        });
                    }
                    finally { ComInterop.Release(button); }
                }

                oleObjects = sheet.OLEObjects;
                dynamic oleObjectsApi = oleObjects;
                for (var index = 1; index <= Convert.ToInt32(oleObjectsApi.Count); index++)
                {
                    object? ole = null;
                    try
                    {
                        ole = oleObjectsApi.Item(index);
                        dynamic oleApi = ole;
                        controls.Add(new
                        {
                            name = Convert.ToString(oleApi.Name),
                            controlType = "activex",
                            controlKind = "activex",
                            left = Convert.ToDouble(oleApi.Left),
                            top = Convert.ToDouble(oleApi.Top),
                            width = Convert.ToDouble(oleApi.Width),
                            height = Convert.ToDouble(oleApi.Height),
                        });
                    }
                    finally { ComInterop.Release(ole); }
                }
                return controls;
            }
            finally
            {
                ComInterop.Release(oleObjects);
                ComInterop.Release(buttons);
            }
        }, sheetName);
    }

    public object CreateForm(JsonElement parameters)
    {
        var formName = parameters.RequiredString("formName");
        var caption = parameters.RequiredString("caption");
        if (!System.Text.RegularExpressions.Regex.IsMatch(formName, "^[A-Za-z_][A-Za-z0-9_]{0,30}$"))
            throw new OfficeWorkerException("invalid_params", "窗体名称必须是 1 到 31 位 VBA 标识符");
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? project = null;
        object? components = null;
        object? form = null;
        object? designer = null;
        object? codeModule = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            project = ((dynamic)workbook).VBProject;
            components = ((dynamic)project).VBComponents;
            dynamic componentsApi = components;
            RemoveNamedComponent(componentsApi, formName, expectedType: 3);
            form = componentsApi.Add(3);
            dynamic formApi = form;
            formApi.Name = formName;
            designer = formApi.Designer;
            dynamic designerApi = designer;
            designerApi.Caption = caption;
            if (parameters.TryGetProperty("width", out var width)) designerApi.Width = width.GetDouble();
            if (parameters.TryGetProperty("height", out var height)) designerApi.Height = height.GetDouble();
            var controls = parameters.PropertyOrEmpty("controls");
            var controlCount = 0;
            if (controls.ValueKind == JsonValueKind.Array)
            {
                foreach (var definition in controls.EnumerateArray())
                {
                    var type = definition.RequiredString("type");
                    var progId = UserFormProgId(type);
                    object? control = null;
                    try
                    {
                        control = designerApi.Controls.Add(progId);
                        dynamic controlApi = control;
                        controlApi.Name = definition.RequiredString("name");
                        controlApi.Left = definition.OptionalDouble("left");
                        controlApi.Top = definition.OptionalDouble("top");
                        controlApi.Width = definition.OptionalDouble("width");
                        controlApi.Height = definition.OptionalDouble("height");
                        var controlCaption = definition.OptionalString("caption");
                        if (!string.IsNullOrEmpty(controlCaption))
                        {
                            try { controlApi.Caption = controlCaption; } catch { }
                        }
                        controlCount++;
                    }
                    finally { ComInterop.Release(control); }
                }
            }
            var eventCode = parameters.OptionalString("eventCode");
            if (!string.IsNullOrWhiteSpace(eventCode))
            {
                codeModule = formApi.CodeModule;
                ((dynamic)codeModule).AddFromString(eventCode);
            }
            return new { success = true, verified = true, formName, caption, controlCount };
        }
        catch
        {
            if (form is not null && components is not null)
            {
                try { ((dynamic)components).Remove(form); } catch { }
            }
            throw;
        }
        finally
        {
            ComInterop.Release(codeModule);
            ComInterop.Release(designer);
            ComInterop.Release(form);
            ComInterop.Release(components);
            ComInterop.Release(project);
            ComInterop.Release(workbook);
        }
    }

    public object AddMenu(JsonElement parameters)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        var menuBar = parameters.RequiredString("menuBar");
        var barName = menuBar switch { "cell" => "Cell", "toolbar" => "Standard", _ => "Worksheet Menu Bar" };
        object? commandBars = null;
        object? commandBar = null;
        object? button = null;
        try
        {
            commandBars = app.CommandBars;
            commandBar = ((dynamic)commandBars).Item(barName);
            button = ((dynamic)commandBar).Controls.Add(1);
            dynamic buttonApi = button;
            buttonApi.Caption = parameters.RequiredString("caption");
            buttonApi.OnAction = parameters.RequiredString("macroName");
            buttonApi.BeginGroup = true;
            var faceId = parameters.OptionalInt32("faceId");
            if (faceId > 0) buttonApi.FaceId = faceId;
            return new { success = true, caption = parameters.RequiredString("caption"), menuBar };
        }
        finally
        {
            ComInterop.Release(button);
            ComInterop.Release(commandBar);
            ComInterop.Release(commandBars);
        }
    }

    private object WithSheet(Func<dynamic, object> operation, string sheetName)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheet = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            sheet = ((dynamic)workbook).Sheets.Item(sheetName);
            return operation(sheet);
        }
        finally
        {
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    private static object AddButton(dynamic sheet, JsonElement parameters, string name, string caption)
    {
        object? buttons = null;
        object? button = null;
        try
        {
            buttons = sheet.Buttons();
            dynamic buttonsApi = buttons;
            try { button = buttonsApi.Item(name); } catch { }
            var created = button is null;
            button ??= buttonsApi.Add(parameters.OptionalDouble("left"), parameters.OptionalDouble("top"), parameters.OptionalDouble("width"), parameters.OptionalDouble("height"));
            dynamic buttonApi = button;
            buttonApi.Name = name;
            buttonApi.Caption = caption;
            var macroName = parameters.OptionalString("macroName");
            if (!string.IsNullOrWhiteSpace(macroName)) buttonApi.OnAction = macroName;
            return new { success = true, created, verified = true, name, controlType = "button", controlKind = "form", caption, onAction = macroName };
        }
        finally
        {
            ComInterop.Release(button);
            ComInterop.Release(buttons);
        }
    }

    private static object AddActiveX(dynamic sheet, JsonElement parameters, string controlType, string name, string caption)
    {
        if (!ControlTypes.TryGetValue(controlType, out var classType))
            throw new OfficeWorkerException("unsupported_control", $"不支持的控件类型: {controlType}");
        if (!string.IsNullOrWhiteSpace(parameters.OptionalString("macroName")))
            throw new OfficeWorkerException("invalid_params", "ActiveX 控件不能通过 OnAction 绑定宏");
        object? oleObjects = null;
        object? ole = null;
        object? control = null;
        try
        {
            oleObjects = sheet.OLEObjects;
            dynamic oleObjectsApi = oleObjects;
            ole = oleObjectsApi.Add(ClassType: classType);
            dynamic oleApi = ole;
            oleApi.Name = name;
            oleApi.Left = parameters.OptionalDouble("left");
            oleApi.Top = parameters.OptionalDouble("top");
            oleApi.Width = parameters.OptionalDouble("width");
            oleApi.Height = parameters.OptionalDouble("height");
            control = oleApi.Object;
            dynamic controlApi = control;
            try { controlApi.Caption = caption; } catch { }
            var linkedCell = parameters.OptionalString("linkedCell");
            if (!string.IsNullOrWhiteSpace(linkedCell)) controlApi.LinkedCell = linkedCell;
            return new { success = true, name, controlType };
        }
        finally
        {
            ComInterop.Release(control);
            ComInterop.Release(ole);
            ComInterop.Release(oleObjects);
        }
    }

    private static void RemoveNamedComponent(dynamic components, string name, int expectedType)
    {
        for (var index = 1; index <= Convert.ToInt32(components.Count); index++)
        {
            object? candidate = components.Item(index);
            try
            {
                dynamic candidateApi = candidate!;
                if (!string.Equals(Convert.ToString(candidateApi.Name), name, StringComparison.OrdinalIgnoreCase)) continue;
                if (Convert.ToInt32(candidateApi.Type) != expectedType)
                    throw new OfficeWorkerException("vba_component_conflict", $"同名 VBA 组件类型不匹配: {name}");
                components.Remove(candidate);
                return;
            }
            finally { ComInterop.Release(candidate); }
        }
    }

    private static string UserFormProgId(string type) => type switch
    {
        "CommandButton" => "Forms.CommandButton.1",
        "TextBox" => "Forms.TextBox.1",
        "Label" => "Forms.Label.1",
        "ComboBox" => "Forms.ComboBox.1",
        "ListBox" => "Forms.ListBox.1",
        "CheckBox" => "Forms.CheckBox.1",
        "OptionButton" => "Forms.OptionButton.1",
        "Frame" => "Forms.Frame.1",
        "SpinButton" => "Forms.SpinButton.1",
        "ScrollBar" => "Forms.ScrollBar.1",
        "Image" => "Forms.Image.1",
        "ToggleButton" => "Forms.ToggleButton.1",
        "TabStrip" => "Forms.TabStrip.1",
        "MultiPage" => "Forms.MultiPage.1",
        _ => throw new OfficeWorkerException("unsupported_control", $"不支持的 UserForm 控件类型: {type}"),
    };
}
