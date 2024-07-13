import zipfile
import os
import re
import datetime
from oletools.olevba import VBA_Parser
from fpdf import FPDF
import graphviz

def extract_vba_code(file_path):
    vba_code = ""
    vba_temp_path = "vbaProject.bin"

    # Ensure the file_path exists
    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        return ""
    
    # Extract the vbaProject.bin file from the .xlsm file
    with zipfile.ZipFile(file_path, 'r') as z:
        for file_info in z.infolist():
            if file_info.filename.endswith('vbaProject.bin'):
                with z.open(file_info) as f:
                    with open(vba_temp_path, 'wb') as temp_file:
                        temp_file.write(f.read())

    # Read the vbaProject.bin file using oletools
    vba_parser = VBA_Parser(vba_temp_path)
    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code_chunk) in vba_parser.extract_macros():
            vba_code += vba_code_chunk

    # Clean up the temporary vbaProject.bin file
    try:
        os.remove(vba_temp_path)
    except PermissionError as e:
        print(f"Failed to remove temporary file {vba_temp_path}: {e}")
    
    return vba_code

def analyze_vba_code(vba_code):
    procedures = re.findall(r'Sub\s+(\w+)', vba_code, re.IGNORECASE)
    functions = re.findall(r'Function\s+(\w+)', vba_code, re.IGNORECASE)
    variables = re.findall(r'Dim\s+(\w+)', vba_code, re.IGNORECASE)
    loops = re.findall(r'(For\s+Each\s+\w+|For\s+\w+\s+in\s+.*|For\s+\w+\s+to\s+.*|Do\s+.*Loop)', vba_code, re.IGNORECASE)
    conditionals = re.findall(r'(If\s+.*Then\s+.*End\s+If|Select\s+Case\s+.*End\s+Select)', vba_code, re.IGNORECASE)
    
    # Perform data flow analysis
    data_flow_analysis = perform_data_flow_analysis(vba_code)

    # Perform code quality analysis
    code_quality = analyze_code_quality(vba_code)

    return {
        "procedures": procedures,
        "functions": functions,
        "variables": variables,
        "loops": loops,
        "conditionals": conditionals,
        "data_flow_analysis": data_flow_analysis,
        "code_quality": code_quality
    }

def perform_data_flow_analysis(vba_code):
    # Placeholder for data flow analysis
    # This function will analyze the flow of data within macros, identify bottlenecks, and suggest optimizations
    bottlenecks = re.findall(r'(Pattern for bottleneck identification)', vba_code, re.IGNORECASE)
    optimizations = re.findall(r'(Pattern for optimization suggestion)', vba_code, re.IGNORECASE)
    return {
        "bottlenecks": bottlenecks,
        "optimizations": optimizations
    }

def analyze_code_quality(vba_code):
    # Placeholder for code quality analysis
    # This function will score the code quality and identify inefficiencies and redundancies
    inefficiencies = re.findall(r'(Pattern for inefficiency)', vba_code, re.IGNORECASE)
    redundancies = re.findall(r'(Pattern for redundancy)', vba_code, re.IGNORECASE)
    score = 100  # Placeholder score calculation
    score -= len(inefficiencies) * 5
    score -= len(redundancies) * 3
    return {
        "score": score,
        "inefficiencies": inefficiencies,
        "redundancies": redundancies,
        "suggestions": ["Optimization suggestion 1", "Optimization suggestion 2"]
    }

def generate_detailed_documentation(vba_code, filename, analysis_results):
    documentation = []
    documentation.append(f"File Name: {filename}")
    documentation.append(f"Date and Time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    documentation.append("\nAnalysis Overview:")
    documentation.append(f"Procedures: {len(analysis_results['procedures'])}")
    documentation.append(f"Functions: {len(analysis_results['functions'])}")
    documentation.append(f"Variables: {len(analysis_results['variables'])}")
    documentation.append(f"Loops: {len(analysis_results['loops'])}")
    documentation.append(f"Conditionals: {len(analysis_results['conditionals'])}")
    documentation.append(f"Code Quality Score: {analysis_results['code_quality']['score']}")
    documentation.append("\nDetailed Analysis:")

    for proc in analysis_results['procedures']:
        documentation.append(f"Procedure '{proc}': This procedure performs specific actions defined within it.")

    for func in analysis_results['functions']:
        documentation.append(f"Function '{func}': This function performs a specific computation or action.")

    for var in analysis_results['variables']:
        documentation.append(f"Variable '{var}': This variable is used to store data.")

    for loop in analysis_results['loops']:
        documentation.append(f"Loop '{loop}': This loop iterates through a set of statements.")

    for cond in analysis_results['conditionals']:
        documentation.append(f"Conditional '{cond}': This conditional determines the flow of execution.")

    # Include code quality analysis in the documentation
    documentation.append("\nCode Quality Analysis:")
    documentation.append(f"Score: {analysis_results['code_quality']['score']}")
    documentation.append("Inefficiencies:")
    for ineff in analysis_results['code_quality']['inefficiencies']:
        documentation.append(f"- {ineff}")
    documentation.append("Redundancies:")
    for red in analysis_results['code_quality']['redundancies']:
        documentation.append(f"- {red}")
    documentation.append("Suggestions:")
    for sugg in analysis_results['code_quality']['suggestions']:
        documentation.append(f"- {sugg}")

    # Include data flow analysis in the documentation
    documentation.append("\nData Flow Analysis:")
    documentation.append("Bottlenecks:")
    for bottleneck in analysis_results['data_flow_analysis']['bottlenecks']:
        documentation.append(f"- {bottleneck}")
    documentation.append("Optimization Opportunities:")
    for optimization in analysis_results['data_flow_analysis']['optimizations']:
        documentation.append(f"- {optimization}")

    return "\n".join(documentation)

def generate_pdf(documentation, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for line in documentation.split('\n'):
        pdf.multi_cell(0, 10, line)

    pdf_output_path = os.path.join('static/pdfs', filename)
    pdf.output(pdf_output_path)
    
    return filename


def generate_flowchart(vba_code):
    dot = graphviz.Digraph(comment='VBA Macro Flow')

    # Simple parsing and node creation logic
    dot.node('Start', 'Start')
    
    procedures = re.findall(r'Sub\s+(\w+)', vba_code, re.IGNORECASE)
    for proc in procedures:
        dot.node(proc, f'Procedure: {proc}')
        dot.edge('Start', proc)

    functions = re.findall(r'Function\s+(\w+)', vba_code, re.IGNORECASE)
    for func in functions:
        dot.node(func, f'Function: {func}')
        dot.edge('Start', func)
    
    dot.node('End', 'End')
    
    for proc in procedures:
        dot.edge(proc, 'End')

    for func in functions:
        dot.edge(func, 'End')

    # Save flowchart as SVG file
    svg_path = os.path.join('static', 'flowcharts', 'flowchart.svg')
    os.makedirs(os.path.dirname(svg_path), exist_ok=True)
    dot.render(filename='flowchart', format='svg', directory='static/flowcharts', cleanup=True)

    return 'flowcharts/flowchart.svg'
