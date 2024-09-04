import os
import win32com.client

# Set paths
reference_file = "reference.SLDRT"
tograde_folder = "/tograde"
graded_folder = "/graded-file"

# Initialize SolidWorks
swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible = True

# Load the reference SLDRT file
reference_doc = swApp.OpenDoc(reference_file, 1)  # 1 is for Part Document

def compare_files(reference_doc, test_doc):
    # Compare properties between reference_doc and test_doc
    # Example: Compare dimensions, features, etc.
    differences = []
    
    # Example comparison logic
    ref_feature_count = reference_doc.FeatureManager.GetFeatureCount()
    test_feature_count = test_doc.FeatureManager.GetFeatureCount()
    
    if ref_feature_count != test_feature_count:
        differences.append(f"Feature count differs: {test_feature_count} vs {ref_feature_count}")
    
    # Add more comparison logic as needed
    return differences

# Iterate over files in the tograde folder
for file_name in os.listdir(tograde_folder):
    if file_name.endswith(".SLDRT"):
        test_file = os.path.join(tograde_folder, file_name)
        
        # Open the test file
        test_doc = swApp.OpenDoc(test_file, 1)
        
        # Compare with reference
        differences = compare_files(reference_doc, test_doc)
        
        # Write differences to a text file in graded-folder
        report_file = os.path.join(graded_folder, file_name.replace(".SLDRT", ".txt"))
        with open(report_file, "w") as f:
            if differences:
                f.write("\n".join(differences))
            else:
                f.write("No differences found")
        
        # Close the test document
        swApp.CloseDoc(test_file)

# Close the reference document
swApp.CloseDoc(reference_file)
