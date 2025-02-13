using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;
using PackagingSolutions.Model;
using System.Windows.Forms;

namespace PackagingSolutions.Services
{
    public class ArasApiService
    {
        private readonly string _serverUrl;
        private readonly string _username;
        private readonly string _database;
        private readonly string _password;
        private Innovator _innovator;

        public ArasApiService(string serverUrl, string database, string username, string password)
        {
            _serverUrl = serverUrl;
            _database = database;
            _username = username;
            _password = password;
        }

        public bool Authenticate()
        {
            try
            {
                HttpServerConnection conn = IomFactory.CreateHttpServerConnection(_serverUrl, _database, _username, _password);
                Item loginResult = conn.Login();
                if (loginResult.isError() || loginResult.isEmpty())
                {
                    ShowError("Login Failed: " + loginResult.getErrorString());
                    return false;
                }

                _innovator = IomFactory.CreateInnovator(conn);
                return true;
            }
            catch (Exception ex)
            {
                ShowError("Error during authentication: " + ex.Message);
                return false;
            }
        }

        public Item? GetPackageByName(string packageName)
        {
            Item package = _innovator.newItem("PackageDefinition", "get");
            package.setProperty("name", packageName);
            Item result = package.apply();

            if (result.isError() || result.isEmpty())
            {
                ShowError($"Package '{packageName}' not found.");
                return null;
            }

            if (result.getItemCount() > 1)
            {
                ShowWarning($"Multiple packages found with name '{packageName}'.");
                return result.getItemByIndex(0);
            }

            return result;
        }

        // Get elements already in the package
        public List<PackageElement> GetElementsInPackage(Item package)
        {
            List<PackageElement> elements = new List<PackageElement>();
            package.fetchRelationships("PackageGroup");
            Item relatedPkgGrp = package.getRelationships("PackageGroup");
            int pkgGrpCount = relatedPkgGrp.getItemCount();
            Item relatedPkgElements = null;

            int pkgElementCount;
            for (int i = 0; i < pkgGrpCount; i++)
            {
                Item pkgGrpItem = relatedPkgGrp.getItemByIndex(i);
                pkgGrpItem.fetchRelationships("PackageElement");
                Item pkgElementItem = pkgGrpItem.getRelationships("PackageElement");
                int elementCount = pkgElementItem?.getItemCount() ?? 0; //Null check

                for (int j = 0; j < elementCount; j++)
                {
                    if (j >= elementCount) break; //Prevent out-of-bounds access
                    Item itm = pkgElementItem.getItemByIndex(j);
               
                    string elementName = itm.getProperty("name");
                    string elementType = itm.getProperty("element_type", "");
                    //ShowWarning(elementName + " " + elementType);
                    elements.Add(new PackageElement
                    {
                        ElementType = elementType,
                        Name = elementName
                    }); ;

                    }

                
            }

            //string message = string.Join("\n", elements);
            //MessageBox.Show(message, "List of Elements", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return elements;
        }

        // Create a new package
        public Item CreatePackage(string packageName)
        {
            Item package = _innovator.newItem("PackageDefinition", "add");
            package.setProperty("name", packageName);
            return package.apply();
        }

        public void AddElementToPackage(Item package, PackageElement element)
        {
            if (package == null || package.isError() || package.isEmpty())
            {
                ShowError("Error: Package is null or invalid.");
                return;
            }

            string packageId = package.getID();
            if (string.IsNullOrEmpty(packageId))
            {
                ShowError("Error: Package ID is empty.");
                return;
            }

            // Declare a local variable to store package elements
            List<PackageElement> existingElements = new List<PackageElement>();

            // Fetch elements only once, if the list is empty
            if (!existingElements.Any())
            {
                existingElements = GetElementsInPackage(package);
            }

            // Check if the element already exists in the package
            bool elementExists = existingElements.Any(e => e.Name == element.Name && e.ElementType == element.ElementType);

            if (elementExists)
            {
                return; // Skip addition if the element is already present
            }

            // Step 2: Fetch the element from Aras
            Item newElement = _innovator.newItem(element.ElementType, "get");
            newElement.setProperty("name", element.Name);
            Item result = newElement.apply();

            if (result.isError() || result.isEmpty())
            {
                ShowError($"Error: Element '{element.Name}' not found in Aras.");
                return;
            }

            if (result.getItemCount() > 1)
            {
                ShowWarning($"Multiple elements found for '{element.Name}'.");
                if (result.getItemCount() > 0) // Ensure items exist before accessing index
                {
                    result = result.getItemByIndex(0);
                }
            }

            string elementId = result.getID();
            //ShowWarning("ElementID: " + elementId);
            if (string.IsNullOrEmpty(elementId))
            {
                ShowError($"Error: Element ID is empty for '{element.Name}'.");
                return;
            }

            Item itemType = _innovator.newItem("ItemType", "get");
            itemType.setProperty("name", element.ElementType);
            itemType= itemType.apply();
            //ShowWarning("element.ElementType: "+element.ElementType);
            string elementLabel = itemType.getProperty("label");
            if (string.IsNullOrEmpty(elementLabel))
            {
                elementLabel = element.ElementType;
            }

            //Step 3: Check if the Package Group already exists
            Item existingPkgGroupQuery = _innovator.newItem("PackageGroup", "get");
            existingPkgGroupQuery.setProperty("source_id", packageId);
            existingPkgGroupQuery.setProperty("name", elementLabel);
            Item existingPkgGroupResult = existingPkgGroupQuery.apply();

            string pkgGrpID;

            if (!existingPkgGroupResult.isError() && !existingPkgGroupResult.isEmpty())
            {
                // Package Group already exists
                pkgGrpID = existingPkgGroupResult.getID();
                
                //ShowSuccess($"Using existing Package Group '{element.ElementType}' for element '{element.Name}'.");
            }
            else
            {
                //Create a new Package Group if it doesn't exist
                Item pkgGrp = _innovator.newItem("PackageGroup", "add");
                pkgGrp.setProperty("source_id", packageId);
                pkgGrp.setProperty("name", elementLabel);
                Item pkgGrpResult = pkgGrp.apply();

                if (pkgGrpResult.isError())
                {
                    ShowError($"Error: Failed to create Package Group '{element.ElementType}'. Reason: {pkgGrpResult.getErrorString()}");
                    return;
                }

                pkgGrpID = pkgGrpResult.getID();
                //ShowSuccess($"Created new Package Group '{element.ElementType}' for element '{element.Name}'.");
            }

            //Step 4: Add element to the existing or new Package Group
            Item existingRel = _innovator.newItem("PackageElement", "add");
            existingRel.setProperty("source_id", pkgGrpID);
            existingRel.setProperty("element_id", elementId);
            existingRel.setProperty("name", element.Name);
            existingRel.setProperty("element_type", element.ElementType);
            Item applyResult = existingRel.apply();

            if (applyResult.isError())
            {
                ShowError($"Error: Failed to add '{element.Name}' to package. Reason: {applyResult.getErrorString()}");
            }
            else
            {
                //ShowSuccess($"Success: '{element.Name}' added to Package Group '{element.ElementType}' in package.");
                //existingElements.Add(element);
            }
        }



        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowSuccess(string message)
        {
            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
