using SolidWorks.Interop.swdocumentmgr;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BlueByte.SOLIDWORKS.DocumentManager.Extensions
{
    public static class Extension
    {
        /// <summary>
        /// Gets the external references and count.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="configurationName">Name of the configuration.</param>
        /// <returns></returns>
        /// <exception cref="Exception">
        /// Failed to get references and count.
        /// or
        /// or
        /// Failed to get references and count. - new Exception("Failed to get configuration names")
        /// or
        /// Failed to get configuration names
        /// or
        /// Failed to get references and count. - new Exception("Specified configuration does not exist in the document.")
        /// or
        /// Specified configuration does not exist in the document.
        /// </exception>
        public static Dictionary<string,int> GetExternalReferencesAndCount(this SwDMDocument doc, bool ignoreSuppressed = true, string configurationName = "")
        {
            var dictionary = new Dictionary<string, int>();

            if (doc.GetVersion() < 2200)
                throw new Exception("Failed to get references and count.", new Exception($"{doc.FullName} is created with SOLIDWORKS 2002 and older. This method does not support older versions of SOLIDWORKS."));
            var configurationNames = doc.GetConfigurationNames();



            if (configurationNames == null)
                throw new Exception("Failed to get references and count.", new Exception("Failed to get configuration names"));


            if (string.IsNullOrWhiteSpace(configurationName))
                configurationName = configurationNames.FirstOrDefault();

            if (configurationNames.Contains(configurationName) == false)
                throw new Exception("Failed to get references and count.", new Exception("Specified configuration does not exist in the document."));

            var config = doc.ConfigurationManager.GetConfigurationByName(configurationName) as SwDMConfiguration3;

            if (config == null)
                throw new Exception("Failed to get references and count.", new Exception("Failed to get document manager configuration."));

            var componentObj = config.GetComponents();

            var components = componentObj as object[];

            if (components != null)
            foreach (var component in components)
            {
                    var swComponent = component as ISwDMComponent11;

                    if (ignoreSuppressed)
                        if (swComponent.IsSuppressed())
                            continue;

                    var fileName = System.IO.Path.GetFileName(swComponent.PathName).ToLower();

                    if (dictionary.ContainsKey(fileName))
                        dictionary[fileName] = dictionary[fileName] + 1;
                    else
                    {
                        dictionary.Add(fileName, 1);
                    }

            }

            return dictionary;

        }

        /// <summary>
        /// Gets the SOLIDWORKS document Manager application.
        /// </summary>
        /// <param name="licenseKey">The license key.</param>
        /// <returns></returns>
        public static SwDMApplication GetApplication(string licenseKey)
        {
            var swClassFact = new SwDMClassFactory();
            return swClassFact.GetApplication(licenseKey);

        }



        /// <summary>
        /// Gets the configuration names.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns></returns>
        public static string[] GetConfigurationNames(this SwDMDocument document)
        {
            return (document.ConfigurationManager.GetConfigurationNames() as object[]).Cast<string>().ToArray();
        }

        /// <summary>
        /// Opens the document.
        /// </summary>
        /// <param name="swDocMgr">The SOLIDWORKS document manager.</param>
        /// <param name="sDocFileName">Name of the s document file.</param>
        /// <param name="openError">The open error.</param>
        /// <param name="openReadOnly">if set to <c>true</c> [open read only].</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">swDocMgr is null.</exception>
        public static SwDMDocument OpenDocument(this SwDMApplication swDocMgr, string sDocFileName, out SwDmDocumentOpenError openError, bool openReadOnly =true)
        {
            openError = SwDmDocumentOpenError.swDmDocumentOpenErrorNone;

            if (swDocMgr == null)
                throw new ArgumentNullException("swDocMgr");

            SwDMDocument swDoc = default(SwDMDocument);
            SwDmDocumentType nDocType = 0;
            SwDmDocumentOpenError nRetVal = 0;
            

            // Determine type of SOLIDWORKS file based on file extension 

            if (sDocFileName.ToLower().EndsWith("sldprt"))
            {
                nDocType = SwDmDocumentType.swDmDocumentPart;
            }
            else if (sDocFileName.ToLower().EndsWith("sldasm"))
            {
                nDocType = SwDmDocumentType.swDmDocumentAssembly;
            }
            else if (sDocFileName.ToLower().EndsWith("slddrw"))
            {
                nDocType = SwDmDocumentType.swDmDocumentDrawing;
            }
            else
            {
                // Not a SOLIDWORKS file 
                nDocType = SwDmDocumentType.swDmDocumentUnknown;
                // So cannot open 
                return null;
            }


            swDoc = swDocMgr.GetDocument(sDocFileName, nDocType, openReadOnly, out nRetVal);

            return swDoc;
        }

        /// <summary>
        /// Closes the document.
        /// </summary>
        /// <param name="swDocMgr">The SOLIDWORKS document manager.</param>
        /// <param name="document">The document.</param>
        /// <exception cref="ArgumentNullException">swDocMgr is null</exception>
        public static void CloseDocument(this SwDMApplication swDocMgr, object document)
        {

            if (swDocMgr == null)
                throw new ArgumentNullException("swDocMgr");
             if (document != null)
            {
                var swDoc = document as ISwDMDocument; 
                if (swDoc != null)
                {
                    swDocMgr.CloseDocument(swDoc);
                }
            }

        }
    }
}
