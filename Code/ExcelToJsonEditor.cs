using UnityEditor;
using UnityEngine;
using System.IO;
using ExcelDataReader;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Text;
using System.Globalization;


public class ExcelToJsonEditor : EditorWindow
{
    private string excelFilePath = string.Empty;
    private string jsonSavePath = string.Empty;
    private string csSavePath = string.Empty;
    private string defaultJsonFileName = "*.json";

    [MenuItem("Tools/Excel To JSON Converter")]
    public static void ShowWindow()
    {
        GetWindow<ExcelToJsonEditor>("Excel to JSON");
    }

    private void OnGUI()
    {
        GUILayout.Label("Ver. 1.1.0", EditorStyles.boldLabel);
        DrawLine();

        var dragArea = GUILayoutUtility.GetRect(0f, 100f, GUILayout.ExpandWidth(true));
        GUI.Box(dragArea, "Drop Excel file here or Click 'Save JSON As'");

        if (dragArea.Contains(Event.current.mousePosition))
        {
            if (Event.current.type == EventType.DragUpdated)
            {
                DragAndDrop.visualMode = DragAndDropVisualMode.Copy;
                Event.current.Use();
            }
            else if (Event.current.type == EventType.DragPerform)
            {
                DragAndDrop.AcceptDrag();

                foreach (var draggedObject in DragAndDrop.objectReferences)
                {
                    var path = AssetDatabase.GetAssetPath(draggedObject);
                    if (!string.IsNullOrEmpty(path) && path.EndsWith(".xlsx"))
                    {
                        excelFilePath = path;

                        // ���ο� .xlsx ������ ���õ� �� ���� ��� �ʱ�ȭ
                        jsonSavePath = "";
                        csSavePath = "";
                        Debug.Log("Excel File Selected: " + path);
                    }
                }
                Event.current.Use();
            }


        }

        if (GUILayout.Button("Save JSON As") && !string.IsNullOrEmpty(excelFilePath))
        {
            string initialPath = Application.dataPath;
            jsonSavePath = EditorUtility.SaveFilePanel("Save JSON as", initialPath, defaultJsonFileName, "json");
            if (!string.IsNullOrEmpty(jsonSavePath))
            {
                ConvertExcelToJson(excelFilePath, jsonSavePath);
                GenerateCSharpClass(excelFilePath, jsonSavePath);
            }
        }
        // �Ǽ� �׸���
        DrawLine();



        GUILayout.BeginHorizontal();
        GUILayout.Label("Selected Excel File:    ", GUILayout.ExpandWidth(false));
        GUI.enabled = false;
        GUILayout.TextField(excelFilePath, GUILayout.ExpandWidth(true));
        GUI.enabled = true;
        GUILayout.EndHorizontal();


        GUILayout.BeginHorizontal();
        GUILayout.Label("JSON Save Path:        ",GUILayout.ExpandWidth(false));
        GUI.enabled = false;
        GUILayout.TextField(jsonSavePath, GUILayout.ExpandWidth(true));
        GUI.enabled = true;
        GUILayout.EndHorizontal();

        GUILayout.BeginHorizontal();
        GUILayout.Label("CS Save Path:             ", GUILayout.ExpandWidth(false));
        GUI.enabled = false;
        GUILayout.TextField(csSavePath, GUILayout.ExpandWidth(true));
        GUI.enabled = true;
        GUILayout.EndHorizontal();
    }

    private void ConvertExcelToJson(string excelPath, string jsonPath)
    {
        using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                var dataTable = result.Tables[0];
                bool isEmptyCellFound = false; // �� �� �߰� ���θ� ����

                List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
                var columnNames = new List<string>();

                //ù ��° ������ ������ �̸� ��������
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    columnNames.Add(dataTable.Rows[0][col].ToString());
                }

                //�� ��° ������ ������ �� ��������
                for (int row = 2; row < dataTable.Rows.Count; row++)
                {
                    var dict = new Dictionary<string, object>();
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        var cellValue = dataTable.Rows[row][col];
                        if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            isEmptyCellFound = true; // �� �� �߰�
                        }
                        var columnName = columnNames[col];
                        var dataType = dataTable.Rows[1][col].ToString();

                        object convertedValue = ConvertCellValue(cellValue, dataType);
                        dict[columnName] = convertedValue;
                    }
                    list.Add(dict);
                }

                string json = JsonConvert.SerializeObject(list, Formatting.Indented);
                File.WriteAllText(jsonPath, json, Encoding.UTF8);
                Debug.Log("Excel to JSON conversion successful: " + jsonPath);

                if (isEmptyCellFound)
                {
                    // �� ���� �߰ߵ� ��� ����ڿ��� �˸�
                    EditorUtility.DisplayDialog("Warning", "Empty or null cell(s) found in the Excel file. Please check your data.", "OK");

                    // ���� �Ϸ� Alert â ǥ��
                    EditorUtility.DisplayDialog("Conversion Complete", "Excel to JSON conversion successful!", "OK");
                }
                else
                {
                    // ���� �Ϸ� Alert â ǥ��
                    EditorUtility.DisplayDialog("Conversion Complete", "Excel to JSON conversion successful!", "OK");
                }
            }
        }
    }


    private object ConvertCellValue(object cellValue, string dataType)
    {
        switch (dataType.ToLower())
        {
            case "int":
                // cellValue�� ����ִ� ���, 0�� ��ȯ
                if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    return 0;
                }
                return int.TryParse(cellValue.ToString(), out int intValue) ? intValue : 0;
            case "float":
                // cellValue�� ����ִ� ���, 0f�� ��ȯ
                if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    return 0f;
                }
                return float.TryParse(cellValue.ToString(), out float floatValue) ? floatValue : 0f;
            case "bool":
                // cellValue�� ����ִ� ���, false�� ��ȯ
                if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    return false;
                }
                return bool.TryParse(cellValue.ToString(), out bool boolValue) ? boolValue : false;
            case "string":
                // cellValue�� ����ִ� ���, null�� ��ȯ
                return cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()) ? null : cellValue.ToString();
            default:
                return null; // �� �� ���� Ÿ���� ��� null ��ȯ
        }
    }


    private void GenerateCSharpClass(string excelPath, string jsonPath)
    {
        string className = ConvertFileNameToClassName(Path.GetFileNameWithoutExtension(jsonPath));
        string scriptPath = Path.Combine(Application.dataPath, "Scripts", "Data"); // �⺻ ��� Assets/Scripts/Data
        if (!Directory.Exists(scriptPath))
        {
            Directory.CreateDirectory(scriptPath);
        }

        string classFilePath = Path.Combine(scriptPath, $"{className}.cs");

        // ������ �̹� �����ϴ��� Ȯ��
        if (File.Exists(classFilePath))
        {
            // ����ڿ��� ��� ������ �����
            bool overwrite = EditorUtility.DisplayDialog("File Exists", $"The file '{className}.cs' already exists. Do you want to overwrite it?", "Yes", "No");

            if (!overwrite)
            {
                // ����ڰ� 'No'�� ������ ��� �۾� ���
                Debug.Log("Operation cancelled by the user.");
                csSavePath = "";
                return;
            }
        }

        using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                var dataTable = result.Tables[0];

                StringBuilder classBuilder = new StringBuilder();
                classBuilder.AppendLine("using System;");
                classBuilder.AppendLine();
                classBuilder.AppendLine($"public class {className}");
                classBuilder.AppendLine("{");

                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    string propertyName = dataTable.Rows[0][col].ToString();
                    string propertyType = ConvertExcelTypeToCSharpType(dataTable.Rows[1][col].ToString());
                    classBuilder.AppendLine($"    public {propertyType} {propertyName};");
                }

                classBuilder.AppendLine("}");
                File.WriteAllText(classFilePath, classBuilder.ToString());
            }
        }
        csSavePath = classFilePath;
        AssetDatabase.Refresh();
        Debug.Log($"C# class '{className}' generated at '{classFilePath}'");
    }

    private string ConvertFileNameToClassName(string fileName)
    {
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(fileName.Replace("_", " ").ToLower()).Replace(" ", string.Empty);
    }

    private string ConvertExcelTypeToCSharpType(string excelType)
    {
        switch (excelType.ToLower())
        {
            case "int":
                return "int";
            case "float":
                return "float";
            case "bool":
                return "bool";
            case "string":
            default:
                return "string";
        }
    }
    private void DrawLine()
    {
        var rect = GUILayoutUtility.GetRect(1, 1, GUILayout.ExpandWidth(true));
        EditorGUI.DrawRect(rect, Color.gray);
    }
}
