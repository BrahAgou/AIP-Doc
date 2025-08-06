from purview_api import get_labels_with_policies
from ppt_generator import create_ppt

def main():
    print("Starting Purview Labels Export...")
    labels_data = get_labels_with_policies()
    create_ppt(labels_data, output_file="PurviewLabelReport.pptx")
    print("PowerPoint report generated: PurviewLabelReport.pptx")

if __name__ == "__main__":
    main()