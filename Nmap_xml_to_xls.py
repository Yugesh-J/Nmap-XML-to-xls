import openpyxl
import xml.etree.ElementTree as ET

def nmap_xml_to_xls(xml_file, xls_file):
    # Parse the XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Add a new worksheet to the workbook
    worksheet = workbook.create_sheet('Sheet1')

    # Write the headers to the Excel worksheet
    worksheet.append(['IP', 'Port', 'Protocol', 'service', 'Status'])

    # Iterate through the hosts in the XML file and write the host information to the Excel worksheet
    for host in root.findall('host'):
        ip = host.find('address').get('addr')
        for port in host.findall('ports/port'):
            port_number = port.get('portid')
            protocol = port.get('protocol')
            try:
                service = port.find('service').get('name')
            except AttributeError:
                service = None
            state = port.find('state').get('state')
            worksheet.append([ip, port_number, protocol, service, state])

    # Save the workbook
    workbook.save(xls_file)

# Test the function
nmap_xml_to_xls('result.xml', 'nmap_result.xlsx')
