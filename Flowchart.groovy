import graphviz
import os

# Create a directed graph
dot = graphviz.Digraph('File Matcher Flowchart', format='png')

# Define nodes
dot.node('A', 'Start')
dot.node('B', 'Upload Statement File')
dot.node('C', 'Upload Estimate File')
dot.node('D', 'Upload Query Results File (Optional)')
dot.node('E', 'Load Statement and Estimate Files')
dot.node('F', 'Check Common Key (PO or ROID)')
dot.node('G', 'Perform Left Join on Common Key')
dot.node('H', 'Update Match Status')
dot.node('I', 'Process Query Results if Uploaded')
dot.node('J', 'Update Match Status Based on Query Results')
dot.node('K', 'Display Matched & Unmatched Records')
dot.node('L', 'Download Matched Results')
dot.node('M', 'End')

# Define edges
edges = [
    ('A', 'B'), ('B', 'C'), ('C', 'D'), ('D', 'E'), ('E', 'F'),
    ('F', 'G'), ('G', 'H'), ('H', 'I'), ('I', 'J'), ('J', 'K'),
    ('K', 'L'), ('L', 'M')
]

for start, end in edges:
    dot.edge(start, end)

# Alternative flows for missing common key or missing files
dot.edge('F', 'M', label='No Common Key Found', style='dashed')
dot.edge('B', 'M', label='File Missing', style='dashed')
dot.edge('C', 'M', label='File Missing', style='dashed')

# Define output directory
output_dir = "C:/Users/Manisha/Downloads"  # Change this to your preferred directory
output_path = os.path.join(output_dir, "file_matcher_flowchart")

# Save and render
dot.render(output_path)

print(f"Flowchart saved at: {output_path}.png")
