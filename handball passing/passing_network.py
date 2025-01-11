import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt

def draw_pass_network(file_path):
    df = pd.read_excel(file_path)

    # checking the data
    if 'From (Player)' not in df.columns or 'To (Player)' not in df.columns or 'Number of Passes' not in df.columns:
        print("file should have 'From (Player)', 'To (Player)', 'Number of Passes' three columns")
        return

    # building the graph
    G = nx.DiGraph()

    # adding links
    for _, row in df.iterrows():
        passer = row['From (Player)']
        receiver = row['To (Player)']
        passes = row['Number of Passes']
        G.add_edge(passer, receiver, weight=passes)

    # calculating indicators
    degree = dict(G.degree())
    in_degree = dict(G.in_degree())
    out_degree = dict(G.out_degree())
    betweenness = nx.betweenness_centrality(G, weight='weight')
    closeness = nx.closeness_centrality(G)
    density = nx.density(G)
    
    # calculating the shortest_path_length 
    if nx.is_connected(G.to_undirected()):
        avg_path_length = nx.average_shortest_path_length(G)
    else:
        avg_path_length = None

    # save indicators to dataframe
    centrality_df = pd.DataFrame({
        'Player': list(G.nodes),
        'Degree': [degree[node] for node in G.nodes],
        'In-Degree': [in_degree[node] for node in G.nodes],
        'Out-Degree': [out_degree[node] for node in G.nodes],
        'Betweenness': [betweenness[node] for node in G.nodes],
        'Closeness': [closeness[node] for node in G.nodes]
    })

    # save indicators to file
    output_path = 'your path'

    centrality_df.to_excel(output_path, index=False)
    
    print(f"indicators save to the {output_path}")

    # draw the figure
    plt.figure(figsize=(16, 12))
    pos = nx.spring_layout(G, seed=42)

    # edge width and color parameters
    edges = G.edges(data=True)
    max_weight = max([data['weight'] for _, _, data in edges])
    edge_widths = [2 + 8 * data['weight'] / max_weight for _, _, data in edges]  
    
    edge_colors = [(0.3 + 0.7 * data['weight'] / max_weight) for _, _, data in edges]  # Minimum opacity 0.3

    # drawing links
    nx.draw_networkx_edges(
        G, pos, edgelist=edges, width=edge_widths, alpha=0.8,
        edge_color=edge_colors, edge_cmap=plt.cm.Greens, edge_vmin=0, edge_vmax=1
    )

    # drawing nodes
    node_sizes = [800 + 100 * degree[node] for node in G.nodes()]
    node_x = [pos[node][0] for node in G.nodes()]
    node_y = [pos[node][1] for node in G.nodes()]
    
    #
    for node in G.nodes():
        # Convert float to int for display
        node_label = str(int(float(node)))
        plt.text(pos[node][0], pos[node][1], node_label,
                fontsize=12, 
                ha='center', 
                va='center', 
                bbox=dict(facecolor='none', edgecolor='none', alpha=0.7),
                zorder=3)
    
    plt.scatter(node_x, node_y, s=node_sizes, c='lightgreen', edgecolors='darkgreen', linewidths=2, alpha=1.0)

    # Set metrics text to top-left corner
    avg_path_length_text = f"{avg_path_length:.2f}" if avg_path_length is not None else "N/A"
    metrics_text = (
        f"Network Density: {density:.2f}\n"
        f"Avg Path Length: {avg_path_length_text}\n"
        f"Nodes: {len(G.nodes)}\n"
        f"Edges: {len(G.edges)}\n"
        f"Avg Degree: {sum(degree.values()) / len(G.nodes):.2f}\n"
        f"Avg Closeness: {sum(closeness.values()) / len(G.nodes):.2f}\n"
        f"Avg Betweenness: {sum(betweenness.values()) / len(G.nodes):.2f}"
    )


    plt.text(0.02, 0.98, metrics_text,
             transform=plt.gca().transAxes,
             fontsize=12,
             bbox=dict(facecolor='white', edgecolor='darkgreen', alpha=0.9),
             verticalalignment='top')

    plt.title("xxxx Passing Network", fontsize=24, fontweight='bold', pad=20)
    plt.axis('off')

    # save figure
    plt.savefig('xxx_passing_network.png',
                format='png', dpi=300, bbox_inches='tight', pad_inches=0.1)

    plt.show()

# file path
file_path = 'your path'
draw_pass_network(file_path)