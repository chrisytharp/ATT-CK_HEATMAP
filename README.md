ATT&CK Heatmap Project

Welcome to the ATT&CK Heatmap Project, designed to enhance your organization's security posture through comprehensive analysis of incident data using the MITRE ATT&CK framework.

Objective:
The goal of this project is to compile lessons learned from incidents into an actionable heatmap, providing insights into the effectiveness of your security controls and identifying areas for improvement.

Requirements:

Python with the openpyxl library installed.

Steps:

Post-Incident Analysis: After each incident, document the Tactics, Techniques, and Procedures (TTPs) discovered using the ATT&CK Navigator Tool available at MITRE ATT&CK Navigator.
Export Data: Save the TTP mappings as a single layer in an Excel (.xlsx) file, renaming it appropriately for future reference.
Execution: Run the provided script.py within the same directory as the TTP mappings Excel file.
Output: The script will generate a Heat_Mappings.xlsx file, presenting the heatmap visualization.

![image](https://github.com/chrisytharp/ATT-CK_HEATMAP/assets/37886152/8ebc5bdd-ae79-4c85-a880-4c424aa4bd20)

Usage:
This project can be executed on a quarterly, bi-annual, or yearly basis to track the success of exploits within your environment and inform adjustments to security controls.

Interpretation:
The Heat_Mappings.xlsx file provides a visual representation of incident data, allowing stakeholders to identify patterns, strengths, and weaknesses in your organization's security posture.
