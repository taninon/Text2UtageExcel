using System.Collections;
using System.Collections.Generic;
using UnityEditor;
using UnityEngine;
using Utage;

[CustomEditor(typeof(Text2UtageExcelConverter))]
public class Text2UtageExcelConverterEditor : Editor
{
	public override void OnInspectorGUI()
	{
		//元のInspector部分を表示
		base.OnInspectorGUI();

		Text2UtageExcelConverter converter = target as Text2UtageExcelConverter;

		if (GUILayout.Button("Convert実行"))
		{
			converter.Convert();
		}
	}
}