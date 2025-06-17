import graphviz
import os

# Create a directed graph
dot = graphviz.Digraph('File Matcher Flowchart', format='png', node_attr={'shape': 'rectangle', 'style': 'filled', 'fillcolor': 'lightblue'})

# Define nodes
nodes = {
    'A': 'Start',
    'B': 'Upload Statement File',
    'C': 'Upload Estimate File',
    'D': 'Upload Query Results File (Optional)',
    'E': 'Load Statement and Estimate Files',
    'F': 'Check Common Key (PO or ROID)',
    'G': 'Perform Left Join on Common Key',
    'H': 'Update Match Status',
    'I': 'Process Query Results if Uploaded',
    'J': 'Update Match Status Based on Query Results',
    'K': 'Display Matched & Unmatched Records',
    'L': 'Download Matched Results',
    'M': 'End'
}

# Add nodes to graph
for key, label in nodes.items():
    dot.node(key, label)

# Define edges for main flow
edges = [
    ('A', 'B'), ('B', 'C'), ('C', 'D'), ('D', 'E'), ('E', 'F'),
    ('F', 'G'), ('G', 'H'), ('H', 'I'), ('I', 'J'), ('J', 'K'),
    ('K', 'L'), ('L', 'M')
]

# Add edges to graph
for start, end in edges:
    dot.edge(start, end)

# Alternative flows for missing common key or missing files
error_edges = [
    ('F', 'M', 'No Common Key Found'),
    ('B', 'M', 'File Missing'),
    ('C', 'M', 'File Missing')
]

for start, end, label in error_edges:
    dot.edge(start, end, label=label, style='dashed', color='red')

# Define output path (change as needed)
output_dir = "C:/Users/Manisha/Downloads"  # Change this to your preferred directory
output_path = os.path.join(output_dir, 'file_matcher_flowchart')

# Save and render
output_file = dot.render(output_path)
print(f"Flowchart saved as: {output_file}")
