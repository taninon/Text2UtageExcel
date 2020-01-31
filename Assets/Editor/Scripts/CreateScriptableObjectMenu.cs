using UnityEngine;
using UnityEditor;

namespace CreateScriptableObjectMenu
{

    /// <summary>
    /// 右クリックメニューからScriptableObjectを作成するエディター拡張
    /// </summary>
    public static class CreateScriptableObjectMenu
    {
        const string MENU_TEXT = "Assets/Create/ScriptableObject";

        [MenuItem(MENU_TEXT, false, 0)]
        static void CreateAsset()
        {
            var script = Selection.activeObject as MonoScript;
            string path = AssetDatabase.GetAssetPath(script);
            Create(script.GetClass(), path.Substring(0, path.Length - 3) + ".asset");
        }

        static void Create(System.Type type, string path)
        {
            ProjectWindowUtil.CreateAsset(ScriptableObject.CreateInstance(type), path);
        }

        [MenuItem(MENU_TEXT, true)]
        static bool ValidateCreateAsset()
        {
            var script = Selection.activeObject as MonoScript;
            if (script == null) { return false; }

            // 選択しているスクリプトがScriptableObjectかどうか
            return script.GetClass().IsSubclassOf(typeof(ScriptableObject));
        }
    }
}