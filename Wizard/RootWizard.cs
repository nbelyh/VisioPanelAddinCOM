
using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Collections.Generic;
using EnvDTE;
using Microsoft.VisualStudio.TemplateWizard;

namespace PanelAddinWizard
{
	/// <summary>
	/// Summary description for SampleWizard.
	/// </summary>
	public class RootWizard 
        : IWizard
	{
        // Use to communicate $saferootprojectname$ to ChildWizard
        public static Dictionary<string, string> GlobalDictionary =
            new Dictionary<string, string>();

        // Add global replacement parameters
        public void RunStarted(object automationObject,
            Dictionary<string, string> replacementsDictionary,
            WizardRunKind runKind, object[] customParams)
        {
            var wizardForm = new WizardForm();

            wizardForm.TaskPane = true;
            wizardForm.Ribbon = true;

            if (wizardForm.ShowDialog() == DialogResult.Cancel)
                throw new WizardBackoutException();

            GlobalDictionary["$csprojectname$"] = replacementsDictionary["$safeprojectname$"];
            GlobalDictionary["$progid$"] = replacementsDictionary["$safeprojectname$"] + ".Addin";
            GlobalDictionary["$clsid$"] = replacementsDictionary["$guid1$"];
            GlobalDictionary["$wixproject$"] = replacementsDictionary["$guid2$"];
            GlobalDictionary["$csprojectguid$"] = replacementsDictionary["$guid3$"];

            GlobalDictionary["$mergeguid$"] = replacementsDictionary["$guid4$"];

            GlobalDictionary["$ribbon$"] = wizardForm.Ribbon ? "true" : "false";
            GlobalDictionary["$commandbars$"] = wizardForm.CommandBars ? "true" : "false";

            GlobalDictionary["$ribbonORcommandbars$"] = wizardForm.Ribbon || wizardForm.CommandBars ? "true" : "false";
            GlobalDictionary["$ribbonANDcommandbars$"] = wizardForm.Ribbon && wizardForm.CommandBars ? "true" : "false";
            GlobalDictionary["$commandbarsANDtaskpane$"] = wizardForm.CommandBars && wizardForm.TaskPane ? "true" : "false";
            GlobalDictionary["$taskpane$"] = wizardForm.TaskPane ? "true" : "false";
            GlobalDictionary["$ui$"] = (wizardForm.CommandBars || wizardForm.Ribbon) ? "true" : "false";
            GlobalDictionary["$taskpaneANDui$"] = wizardForm.TaskPane && (wizardForm.CommandBars || wizardForm.Ribbon) ? "true" : "false";
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }
    }
}
