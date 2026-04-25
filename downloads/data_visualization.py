import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import os
from datetime import datetime

# Set font to support special characters
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['Arial']
matplotlib.rcParams['axes.unicode_minus'] = False

# ================= CONFIGURATION =================
TRACKING_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tracking Data")

def analyze_existing_data():
    """Find and analyze the latest Excel file."""
    
    # Find latest Excel file
    excel_files = [f for f in os.listdir(TRACKING_DATA_DIR) if f.endswith('.xlsx') and 'Analysis' in f]
    
    if not excel_files:
        print("❌ No tracking analysis Excel file found!")
        print(f"📂 Looking in: {TRACKING_DATA_DIR}")
        return None
    
    # Get latest file
    latest_file = max(excel_files, key=lambda x: os.path.getctime(os.path.join(TRACKING_DATA_DIR, x)))
    excel_path = os.path.join(TRACKING_DATA_DIR, latest_file)
    
    print(f"📊 Analyzing: {latest_file}")
    print("="*50)
    
    # Read data
    df = pd.read_excel(excel_path, sheet_name='Tracking Data')
    
    return df, latest_file


def create_visualizations(df, source_file):
    """Create multiple visualization graphs."""
    
    # Create output directory for graphs
    graphs_dir = os.path.join(TRACKING_DATA_DIR, "Graphs")
    os.makedirs(graphs_dir, exist_ok=True)
    
    # Set style
    plt.style.use('seaborn-v0_8-darkgrid')
    
    # ================= GRAPH 1: Timely vs Late Delivery (Pie Chart) =================
    fig1, ax1 = plt.subplots(figsize=(10, 8))
    
    # Count delayed vs on-time
    delayed_count = len(df[df['is_delayed'] == 'Yes'])
    ontime_count = len(df[df['is_delayed'] == 'No'])
    transit_count = len(df[df['delivery_status'] == 'In Transit'])
    
    colors = ['#ff6b6b', '#51cf66', '#4dabf7']  # Red, Green, Blue
    
    if delayed_count > 0 or ontime_count > 0:
        sizes = [ontime_count, delayed_count]
        labels = [f'On-Time\n({ontime_count})', f'Delayed\n({delayed_count})']
        explode = (0.05, 0.05)
        
        wedges, texts, autotexts = ax1.pie(sizes, explode=explode, labels=labels, 
                                            colors=colors[:2], autopct='%1.1f%%',
                                            shadow=True, startangle=90, textprops={'fontsize': 12})
        
        # Enhance percentage text
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        ax1.set_title('Delivery Performance: On-Time vs Delayed', fontsize=16, fontweight='bold', pad=20)
        
        # Save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        graph1_path = os.path.join(graphs_dir, f"Delivery_Performance_{timestamp}.png")
        plt.savefig(graph1_path, dpi=300, bbox_inches='tight')
        print(f"✅ Saved: {graph1_path}")
        plt.close()
    else:
        print("⚠️  No completed deliveries to analyze for delays")
    
    # ================= GRAPH 2: Delivery Status Breakdown (Bar Chart) =================
    fig2, ax2 = plt.subplots(figsize=(12, 6))
    
    status_counts = df['delivery_status'].value_counts()
    colors_bar = ['#51cf66' if s == 'Delivered' else '#ffd43b' for s in status_counts.index]
    
    bars = ax2.bar(range(len(status_counts)), status_counts.values, color=colors_bar, 
                   edgecolor='black', linewidth=1.5, alpha=0.8)
    
    ax2.set_xticks(range(len(status_counts)))
    ax2.set_xticklabels(status_counts.index, fontsize=12, fontweight='bold')
    ax2.set_ylabel('Number of Packages', fontsize=12, fontweight='bold')
    ax2.set_title('Package Delivery Status Overview', fontsize=16, fontweight='bold', pad=20)
    ax2.grid(axis='y', alpha=0.3)
    
    # Add value labels on bars
    for bar, count in zip(bars, status_counts.values):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1, 
                str(count), ha='center', va='bottom', fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    graph2_path = os.path.join(graphs_dir, f"Status_Overview_{timestamp}.png")
    plt.savefig(graph2_path, dpi=300, bbox_inches='tight')
    print(f"✅ Saved: {graph2_path}")
    plt.close()
    
    # ================= GRAPH 3: Delivery Time Distribution (Histogram) =================
    fig3, ax3 = plt.subplots(figsize=(12, 6))
    
    # Filter valid delivery days
    df_valid = df[df['total_days'] != 'N/A'].copy()
    if not df_valid.empty:
        df_valid['total_days'] = pd.to_numeric(df_valid['total_days'], errors='coerce')
        df_valid = df_valid.dropna(subset=['total_days'])
        
        if not df_valid.empty:
            # Create bins
            bins = range(0, int(df_valid['total_days'].max()) + 2)
            
            ax3.hist(df_valid['total_days'], bins=bins, color='#4dabf7', edgecolor='black', 
                    alpha=0.7, linewidth=1)
            
            # Add threshold line at 7 days
            ax3.axvline(x=7, color='red', linestyle='--', linewidth=2, label='Delay Threshold (7 days)')
            
            ax3.set_xlabel('Delivery Time (Days)', fontsize=12, fontweight='bold')
            ax3.set_ylabel('Number of Packages', fontsize=12, fontweight='bold')
            ax3.set_title('Distribution of Delivery Times', fontsize=16, fontweight='bold', pad=20)
            ax3.legend(fontsize=10)
            ax3.grid(axis='y', alpha=0.3)
            
            graph3_path = os.path.join(graphs_dir, f"Delivery_Time_Distribution_{timestamp}.png")
            plt.savefig(graph3_path, dpi=300, bbox_inches='tight')
            print(f"✅ Saved: {graph3_path}")
            plt.close()
    
    # ================= GRAPH 4: City-wise Delivery Performance (Horizontal Bar) =================
    fig4, ax4 = plt.subplots(figsize=(12, 8))
    
    if 'receive_city' in df.columns:
        city_performance = df.groupby('receive_city').agg({
            'tracking_number': 'count',
            'is_delayed': lambda x: (x == 'Yes').sum()
        }).reset_index()
        city_performance.columns = ['City', 'Total', 'Delayed']
        city_performance['On_Time'] = city_performance['Total'] - city_performance['Delayed']
        
        # Sort by total
        city_performance = city_performance.sort_values('Total', ascending=True)
        
        y_pos = range(len(city_performance))
        
        ax4.barh(y_pos, city_performance['On_Time'], color='#51cf66', label='On-Time', edgecolor='black')
        ax4.barh(y_pos, city_performance['Delayed'], left=city_performance['On_Time'], 
                color='#ff6b6b', label='Delayed', edgecolor='black')
        
        ax4.set_yticks(y_pos)
        ax4.set_yticklabels(city_performance['City'], fontsize=10)
        ax4.set_xlabel('Number of Packages', fontsize=12, fontweight='bold')
        ax4.set_title('Delivery Performance by City', fontsize=16, fontweight='bold', pad=20)
        ax4.legend(loc='lower right', fontsize=10)
        ax4.grid(axis='x', alpha=0.3)
        
        graph4_path = os.path.join(graphs_dir, f"City_Performance_{timestamp}.png")
        plt.savefig(graph4_path, dpi=300, bbox_inches='tight')
        print(f"✅ Saved: {graph4_path}")
        plt.close()
    
    # ================= PRINT SUMMARY =================
    print("\n" + "="*50)
    print("📊 DELIVERY ANALYSIS SUMMARY")
    print("="*50)
    print(f"📦 Total Packages: {len(df)}")
    print(f"✅ Delivered: {len(df[df['delivery_status'] == 'Delivered'])}")
    print(f"🚚 In Transit: {len(df[df['delivery_status'] == 'In Transit'])}")
    print(f"✓ On-Time: {ontime_count}")
    print(f"⚠️  Delayed: {delayed_count}")
    
    if delayed_count + ontime_count > 0:
        ontime_pct = (ontime_count / (delayed_count + ontime_count)) * 100
        print(f"📈 On-Time Rate: {ontime_pct:.1f}%")
    
    print(f"\n📂 Graphs saved in: {graphs_dir}")
    print("="*50)
    
    # Show graphs
    print("\n💡 To view graphs, open the PNG files in the 'Graphs' folder")
    print("   Or run: start " + graphs_dir.replace("\\", "/"))


# ================= RUN ANALYSIS =================
if __name__ == "__main__":
    print("📈 BuffaloEx Delivery Analytics")
    print("="*50)
    
    result = analyze_existing_data()
    
    if result:
        df, source_file = result
        create_visualizations(df, source_file)