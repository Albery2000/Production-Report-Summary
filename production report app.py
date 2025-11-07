import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io
import base64

def extract_wells_with_net_diff_bo(file_content):
    """
    Extract wells that have Net Diff BO values (excluding zeros) from specific columns and stop at TOTAL row
    """
    try:
        # Read the Excel file with multi-level headers, skipping first 6 rows
        df = pd.read_excel(
            file_content, 
            sheet_name='Report', 
            skiprows=6,
            header=[0, 1]  # Two header rows
        )
        
        st.subheader("🔍 Detected Column Structure")
        
        # Display all columns to help with debugging
        columns_info = []
        for i, col in enumerate(df.columns):
            col_info = {
                'Column Index': i,
                'Level 0': str(col[0]) if pd.notna(col[0]) else '',
                'Level 1': str(col[1]) if len(col) > 1 and pd.notna(col[1]) else '',
                'Full Name': str(col)
            }
            columns_info.append(col_info)
        
        columns_df = pd.DataFrame(columns_info)
        st.dataframe(columns_df)
        
        # Find the specific columns we need
        field_col = None
        well_name_col = None
        net_diff_bo_col = None
        net_bo_col = None
        
        for i, col in enumerate(df.columns):
            # Look for the exact column structures
            if str(col) == "('TOTAL PRODUCTION', 'Net diff. BO')":
                net_diff_bo_col = col
                st.success(f"✅ Found Net Diff BO column: {col} (Index {i})")
            
            elif str(col) == "('TOTAL PRODUCTION', 'Net\\nBO')" or "('TOTAL PRODUCTION', 'Net\nBO')" in str(col):
                net_bo_col = col
                st.success(f"✅ Found Net BO column: {col} (Index {i})")
            
            # Field column - look for ('Field', 'Unnamed: 0_level_1')
            elif str(col) == "('Field', 'Unnamed: 0_level_1')":
                field_col = col
                st.success(f"✅ Found Field column: {col} (Index {i})")
            
            # Well name column - look for ('RUNNING WELLS', 'Unnamed: 1_level_1')
            elif str(col) == "('RUNNING WELLS', 'Unnamed: 1_level_1')":
                well_name_col = col
                st.success(f"✅ Found Well Name column: {col} (Index {i})")
        
        # Validation
        if field_col is None:
            st.error("❌ Could not find 'Field' column")
            return None, None, None, None, None
        
        if well_name_col is None:
            st.error("❌ Could not find well name column")
            return None, None, None, None, None
        
        if net_diff_bo_col is None:
            st.error("❌ Could not find 'Net diff. BO' column")
            return None, None, None, None, None
        
        if net_bo_col is None:
            st.error("❌ Could not find 'Net BO' column")
            return None, None, None, None, None
        
        # Convert numeric columns
        df[net_diff_bo_col] = pd.to_numeric(df[net_diff_bo_col], errors='coerce')
        df[net_bo_col] = pd.to_numeric(df[net_bo_col], errors='coerce')
        
        # Find where to stop (at "TOTAL" in Field column)
        stop_index = None
        for idx, value in enumerate(df[field_col]):
            if pd.notna(value) and 'TOTAL' in str(value).upper():
                stop_index = idx
                st.info(f"🛑 Found 'TOTAL' row at index {idx}, stopping extraction here")
                break
        
        # If no TOTAL found, use all rows
        if stop_index is None:
            stop_index = len(df)
            st.warning("⚠️ No 'TOTAL' row found, using all available data")
        
        # Filter rows up to the TOTAL row
        df_before_total = df.iloc[:stop_index].copy()
        
        # Calculate TOTAL statistics for ALL wells (including zeros)
        all_wells_count = len(df_before_total)
        total_net_bo_all = df_before_total[net_bo_col].sum()
        total_net_diff_bo_all = df_before_total[net_diff_bo_col].sum()
        
        # Filter rows that have Net diff. BO values AND are not zero (but include negative values)
        filtered_df = df_before_total[
            (df_before_total[net_diff_bo_col].notna()) & 
            (df_before_total[net_diff_bo_col] != 0)  # Exclude zeros but include negatives
        ].copy()
        
        if filtered_df.empty:
            st.warning("⚠️ No wells found with non-zero Net Diff BO values before TOTAL row")
            return None, None, None, None, None
        
        # Show how many wells were filtered out due to zero values
        all_wells_with_net_diff = df_before_total[df_before_total[net_diff_bo_col].notna()]
        zero_wells_count = len(all_wells_with_net_diff[all_wells_with_net_diff[net_diff_bo_col] == 0])
        st.info(f"📊 Filtered out {zero_wells_count} wells with zero Net Diff BO values")
        
        # Show distribution of positive vs negative values
        positive_count = len(filtered_df[filtered_df[net_diff_bo_col] > 0])
        negative_count = len(filtered_df[filtered_df[net_diff_bo_col] < 0])
        st.info(f"📈 Value distribution: {positive_count} positive, {negative_count} negative Net Diff BO values")
        
        # Select the columns we need in the correct order
        result_columns = [field_col, well_name_col, net_bo_col, net_diff_bo_col]
        
        # Create final result dataframe
        result_df = filtered_df[result_columns].copy()
        
        # Clean up the data - remove rows where well name is empty or is a field name
        field_names = ['Ferdaus', 'Sidra', 'Ganna', 'Rayan', 'Abrar', 'Abrar-South', 'Rawda']
        
        # Filter out rows where well_name is actually a field name
        mask = ~result_df[well_name_col].isin(field_names)
        result_df = result_df[mask].copy()
        
        # Remove rows where well_name is empty or NaN
        result_df = result_df[result_df[well_name_col].notna()]
        result_df = result_df[result_df[well_name_col] != '']
        
        result_df = result_df.reset_index(drop=True)
        
        # Calculate totals and statistics for non-zero wells
        total_net_bo_non_zero = result_df[net_bo_col].sum()
        total_net_diff_bo_non_zero = result_df[net_diff_bo_col].sum()
        well_count_non_zero = len(result_df)
        
        # Calculate statistics for both ALL wells and non-zero wells
        stats = {
            # All Wells Statistics
            'Total All Wells': all_wells_count,
            'Total Net BO (All Wells)': total_net_bo_all,
            'Total Net Diff BO (All Wells)': total_net_diff_bo_all,
            'Average Net BO (All Wells)': df_before_total[net_bo_col].mean(),
            'Average Net Diff BO (All Wells)': df_before_total[net_diff_bo_col].mean(),
            
            # Non-Zero Wells Statistics
            'Total Wells with Non-Zero Net Diff BO': well_count_non_zero,
            'Positive Net Diff BO Wells': positive_count,
            'Negative Net Diff BO Wells': negative_count,
            'Total Net BO (Non-Zero Wells)': total_net_bo_non_zero,
            'Total Net Diff BO (Non-Zero Wells)': total_net_diff_bo_non_zero,
            'Average Net BO (Non-Zero Wells)': result_df[net_bo_col].mean(),
            'Average Net Diff BO (Non-Zero Wells)': result_df[net_diff_bo_col].mean(),
            'Maximum Net BO': result_df[net_bo_col].max(),
            'Maximum Net Diff BO': result_df[net_diff_bo_col].max(),
            'Minimum Net BO': result_df[net_bo_col].min(),
            'Minimum Net Diff BO': result_df[net_diff_bo_col].min(),
            'Median Net BO': result_df[net_bo_col].median(),
            'Median Net Diff BO': result_df[net_diff_bo_col].median(),
            'Standard Deviation Net BO': result_df[net_bo_col].std(),
            'Standard Deviation Net Diff BO': result_df[net_diff_bo_col].std()
        }
        
        # Create the final dataframe with proper column structure
        final_df = result_df.copy()
        
        # Format numeric columns
        for col in [net_bo_col, net_diff_bo_col]:
            if col in final_df.columns and final_df[col].dtype in [np.float64, np.int64]:
                final_df[col] = final_df[col].round(2)
        
        # Add total rows for both ALL wells and non-zero wells
        total_row_non_zero = pd.DataFrame([{
            field_col: 'TOTAL (Non-Zero Wells)',
            well_name_col: f'{well_count_non_zero} Wells with Non-Zero Net Diff BO',
            net_bo_col: total_net_bo_non_zero,
            net_diff_bo_col: total_net_diff_bo_non_zero
        }])
        
        total_row_all = pd.DataFrame([{
            field_col: 'TOTAL (All Wells)',
            well_name_col: f'{all_wells_count} Total Wells',
            net_bo_col: total_net_bo_all,
            net_diff_bo_col: total_net_diff_bo_all
        }])
        
        # Combine main data with total rows
        final_df = pd.concat([final_df, total_row_non_zero, total_row_all], ignore_index=True)
        
        st.success(f"✅ Successfully extracted {well_count_non_zero} wells with non-zero Net Diff BO values")
        
        return final_df, well_count_non_zero, stats, [field_col, well_name_col, net_bo_col, net_diff_bo_col], df_before_total
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None, None, None, None, None

def create_visualizations(data_without_total, original_columns, all_wells_data):
    """
    Create simplified statistical visualizations
    """
    try:
        # Check if we have valid data for visualizations
        if data_without_total.empty or all_wells_data.empty:
            st.warning("No data available for visualizations")
            return None
            
        # Extract the column names
        field_col = original_columns[0]      # ('Field', 'Unnamed: 0_level_1')
        well_name_col = original_columns[1]  # ('RUNNING WELLS', 'Unnamed: 1_level_1')
        net_bo_col = original_columns[2]     # ('TOTAL PRODUCTION', 'Net\nBO')
        net_diff_bo_col = original_columns[3] # ('TOTAL PRODUCTION', 'Net diff. BO')
        
        # Create clean copies for visualization
        viz_data_non_zero = data_without_total.copy()
        viz_data_all = all_wells_data.copy()
        
        # Remove rows with NaN values in the key columns for visualization
        viz_data_non_zero = viz_data_non_zero[
            viz_data_non_zero[well_name_col].notna() & 
            viz_data_non_zero[net_bo_col].notna() & 
            viz_data_non_zero[net_diff_bo_col].notna()
        ]
        
        viz_data_all = viz_data_all[
            viz_data_all[well_name_col].notna() & 
            viz_data_all[net_bo_col].notna()
        ]
        
        # Check if we have any data left after cleaning
        if viz_data_non_zero.empty or viz_data_all.empty:
            st.warning("No valid data available for visualizations after removing NaN values")
            return None
        
        # Extract clean data for visualization
        # Non-zero wells data
        well_names_non_zero = viz_data_non_zero[well_name_col]
        net_bo_data_non_zero = viz_data_non_zero[net_bo_col]
        net_diff_bo_data_non_zero = viz_data_non_zero[net_diff_bo_col]
        
        # All wells data
        well_names_all = viz_data_all[well_name_col]
        net_bo_data_all = viz_data_all[net_bo_col]
        
        # Check for finite values
        if (net_bo_data_non_zero.isna().all() or net_diff_bo_data_non_zero.isna().all() or 
            not np.isfinite(net_bo_data_non_zero).any() or not np.isfinite(net_diff_bo_data_non_zero).any() or
            net_bo_data_all.isna().all() or not np.isfinite(net_bo_data_all).any()):
            st.warning("No finite values available for visualization")
            return None
        
        # Create simplified subplots - 2 rows, 2 columns
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle('Production Analysis Dashboard', fontsize=16, fontweight='bold')
        
        # 1. Net Diff BO by Well (Non-Zero Wells - Top 15)
        if len(net_diff_bo_data_non_zero) > 0 and len(well_names_non_zero) > 0:
            display_data = pd.DataFrame({
                'well_name': well_names_non_zero,
                'net_diff_bo': net_diff_bo_data_non_zero
            }).head(15)
            
            display_wells = display_data['well_name']
            display_net_diff = display_data['net_diff_bo']
            
            bars = axes[0, 0].bar(range(len(display_wells)), display_net_diff, 
                                 color=['lightgreen' if x >= 0 else 'lightcoral' for x in display_net_diff],
                                 alpha=0.7)
            axes[0, 0].set_xlabel('Wells')
            axes[0, 0].set_ylabel('Net Diff BO')
            axes[0, 0].set_title('Net Diff BO Performance (Top 15 Wells)')
            axes[0, 0].set_xticks(range(len(display_wells)))
            axes[0, 0].set_xticklabels(display_wells, rotation=45, ha='right')
            axes[0, 0].grid(True, alpha=0.3)
            
            for bar, value in zip(bars, display_net_diff):
                height = bar.get_height()
                axes[0, 0].text(bar.get_x() + bar.get_width()/2., height,
                               f'{value:.1f}', ha='center', va='bottom' if height >= 0 else 'top',
                               fontsize=8)
        else:
            axes[0, 0].text(0.5, 0.5, 'No data available', ha='center', va='center', transform=axes[0, 0].transAxes)
            axes[0, 0].set_title('Net Diff BO Performance')
        
        # 2. Net BO by Well (Non-Zero Wells - Top 15)
        if len(net_bo_data_non_zero) > 0 and len(well_names_non_zero) > 0:
            display_data = pd.DataFrame({
                'well_name': well_names_non_zero,
                'net_bo': net_bo_data_non_zero
            }).head(15)
            
            display_wells = display_data['well_name']
            display_net_bo = display_data['net_bo']
            
            bars = axes[0, 1].bar(range(len(display_wells)), display_net_bo, 
                                 color='skyblue', alpha=0.7)
            axes[0, 1].set_xlabel('Wells')
            axes[0, 1].set_ylabel('Net BO')
            axes[0, 1].set_title('Net BO Production (Top 15 Wells)')
            axes[0, 1].set_xticks(range(len(display_wells)))
            axes[0, 1].set_xticklabels(display_wells, rotation=45, ha='right')
            axes[0, 1].grid(True, alpha=0.3)
            
            for bar, value in zip(bars, display_net_bo):
                height = bar.get_height()
                axes[0, 1].text(bar.get_x() + bar.get_width()/2., height,
                               f'{value:.0f}', ha='center', va='bottom',
                               fontsize=8)
        else:
            axes[0, 1].text(0.5, 0.5, 'No data available', ha='center', va='center', transform=axes[0, 1].transAxes)
            axes[0, 1].set_title('Net BO Production')
        
        # 3. Top 10 Wells with Highest Net BO (ALL WELLS)
        if len(net_bo_data_all) > 0 and len(well_names_all) > 0:
            # Get top 10 wells with highest Net BO from ALL wells
            top_wells_all = pd.DataFrame({
                'well_name': well_names_all,
                'net_bo': net_bo_data_all
            }).nlargest(10, 'net_bo')
            
            # Create horizontal bar chart for better readability
            bars = axes[1, 0].barh(range(len(top_wells_all)), top_wells_all['net_bo'], 
                                  color='gold', alpha=0.7, edgecolor='darkorange', linewidth=1)
            axes[1, 0].set_xlabel('Net BO')
            axes[1, 0].set_ylabel('Wells')
            axes[1, 0].set_title('Top 10 Highest Producing Wells')
            axes[1, 0].set_yticks(range(len(top_wells_all)))
            axes[1, 0].set_yticklabels(top_wells_all['well_name'])
            axes[1, 0].grid(True, alpha=0.3)
            
            # Add value labels on bars
            for bar, value in zip(bars, top_wells_all['net_bo']):
                width = bar.get_width()
                axes[1, 0].text(width + width*0.01, bar.get_y() + bar.get_height()/2.,
                               f'{value:.0f}', ha='left', va='center', fontsize=9, fontweight='bold')
        else:
            axes[1, 0].text(0.5, 0.5, 'No data available', ha='center', va='center', transform=axes[1, 0].transAxes)
            axes[1, 0].set_title('Top 10 Highest Producing Wells')
        
        # 4. Performance Overview Pie Chart
        if len(net_diff_bo_data_non_zero) > 0:
            # Calculate performance categories
            excellent = len(net_diff_bo_data_non_zero[net_diff_bo_data_non_zero > 100])
            good = len(net_diff_bo_data_non_zero[(net_diff_bo_data_non_zero > 0) & (net_diff_bo_data_non_zero <= 100)])
            poor = len(net_diff_bo_data_non_zero[(net_diff_bo_data_non_zero < 0) & (net_diff_bo_data_non_zero >= -100)])
            critical = len(net_diff_bo_data_non_zero[net_diff_bo_data_non_zero < -100])
            
            categories = ['Excellent (>100)', 'Good (0-100)', 'Poor (-100-0)', 'Critical (<-100)']
            values = [excellent, good, poor, critical]
            colors = ['#2ecc71', '#3498db', '#f39c12', '#e74c3c']
            
            # Only show categories with values
            valid_categories = []
            valid_values = []
            valid_colors = []
            
            for cat, val, col in zip(categories, values, colors):
                if val > 0:
                    valid_categories.append(cat)
                    valid_values.append(val)
                    valid_colors.append(col)
            
            if len(valid_values) > 0:
                wedges, texts, autotexts = axes[1, 1].pie(valid_values, labels=valid_categories, colors=valid_colors, 
                                                         autopct='%1.1f%%', startangle=90)
                axes[1, 1].set_title('Well Performance Distribution')
                
                # Improve readability
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontweight('bold')
            else:
                axes[1, 1].text(0.5, 0.5, 'No performance data', ha='center', va='center', transform=axes[1, 1].transAxes)
                axes[1, 1].set_title('Well Performance Distribution')
        else:
            axes[1, 1].text(0.5, 0.5, 'No data available', ha='center', va='center', transform=axes[1, 1].transAxes)
            axes[1, 1].set_title('Well Performance Distribution')
        
        plt.tight_layout()
        return fig
        
    except Exception as e:
        st.error(f"❌ Error creating visualizations: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None

def create_comprehensive_powerpoint(data_df, well_count, stats, original_columns, visualization_fig):
    """
    Create a comprehensive PowerPoint presentation with data, statistics, and visualizations
    """
    try:
        # Create a new presentation
        prs = Presentation()
        
        # Title slide
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Production Analysis Report"
        subtitle.text = f"Comprehensive Well Performance Analysis\nTotal Wells: {stats['Total All Wells']}\nGenerated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}"
        
        # Executive Summary Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Executive Summary"
        
        # Add summary content
        content_left = Inches(0.5)
        content_top = Inches(1.5)
        content_width = Inches(9.0)
        content_height = Inches(5.0)
        
        text_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        # Add summary points
        summary_points = [
            f"• Total Wells Analyzed: {stats['Total All Wells']}",
            f"• Wells with Non-Zero Net Diff BO: {stats['Total Wells with Non-Zero Net Diff BO']}",
            f"• Positive Performance Wells: {stats['Positive Net Diff BO Wells']}",
            f"• Wells Requiring Attention: {stats['Negative Net Diff BO Wells']}",
            f"• Total Net BO Production: {stats['Total Net BO (All Wells)']:,.0f}",
            f"• Average Net BO per Well: {stats['Average Net BO (All Wells)']:,.0f}",
            f"• Highest Producing Well: {stats['Maximum Net BO']:,.0f}",
            f"• Performance Range: {stats['Minimum Net BO']:,.0f} to {stats['Maximum Net BO']:,.0f}"
        ]
        
        for point in summary_points:
            p = text_frame.add_paragraph()
            p.text = point
            p.space_after = Inches(0.05)
        
        # Main Data Table Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Production Data - Key Wells"
        
        # Create main data table (show only first 15 rows for readability)
        display_data = data_df.head(15) if len(data_df) > 15 else data_df
        
        rows = len(display_data) + 1
        cols = len(display_data.columns)
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9.0)
        height = Inches(0.8 * min(rows, 12))  # Limit height
        
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        
        # Set column headers
        for i, column in enumerate(display_data.columns):
            table.cell(0, i).text = str(column)
        
        # Fill table with data
        for row_idx, (_, row_data) in enumerate(display_data.iterrows(), 1):
            for col_idx, column in enumerate(display_data.columns):
                value = row_data[column]
                if isinstance(value, (int, float)) and column not in [original_columns[0], original_columns[1]]:
                    table.cell(row_idx, col_idx).text = f"{value:,.2f}"
                else:
                    table.cell(row_idx, col_idx).text = str(value)
        
        # Key Metrics Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Key Performance Metrics"
        
        # Create key metrics table
        key_metrics = {
            'Total Wells': stats['Total All Wells'],
            'Wells with Significant Changes': stats['Total Wells with Non-Zero Net Diff BO'],
            'Positive Performance Wells': stats['Positive Net Diff BO Wells'],
            'Wells Requiring Attention': stats['Negative Net Diff BO Wells'],
            'Total Net BO Production': stats['Total Net BO (All Wells)'],
            'Total Net Diff BO': stats['Total Net Diff BO (All Wells)'],
            'Average Net BO per Well': stats['Average Net BO (All Wells)'],
            'Highest Producing Well': stats['Maximum Net BO'],
            'Performance Standard Deviation': stats['Standard Deviation Net BO']
        }
        
        stats_rows = len(key_metrics) + 1
        stats_cols = 2
        left = Inches(1.0)
        top = Inches(1.5)
        width = Inches(8.0)
        height = Inches(0.8 * min(stats_rows, 15))
        
        stats_table = slide.shapes.add_table(stats_rows, stats_cols, left, top, width, height).table
        stats_table.cell(0, 0).text = "Metric"
        stats_table.cell(0, 1).text = "Value"
        
        for idx, (metric, value) in enumerate(key_metrics.items(), 1):
            stats_table.cell(idx, 0).text = metric
            if isinstance(value, (int, float)):
                if value > 1000:
                    stats_table.cell(idx, 1).text = f"{value:,.0f}"
                else:
                    stats_table.cell(idx, 1).text = f"{value:,.2f}"
            else:
                stats_table.cell(idx, 1).text = str(value)
        
        # Visualization Slides
        if visualization_fig:
            # Save figure to bytes
            img_buffer = io.BytesIO()
            visualization_fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
            img_buffer.seek(0)
            
            # Create individual visualization slides
            visualization_titles = [
                "Net Diff BO Performance",
                "Net BO Production", 
                "Top 10 Highest Producing Wells",
                "Well Performance Distribution"
            ]
            
            for viz_title in visualization_titles:
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                title = slide.shapes.title
                title.text = f"Analysis - {viz_title}"
                
                # Add the visualization image
                left = Inches(1.0)
                top = Inches(1.5)
                width = Inches(8.0)
                slide.shapes.add_picture(img_buffer, left, top, width=width)
        
        # Recommendations Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Recommendations & Next Steps"
        
        text_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        recommendations = [
            "🎯 Focus Areas:",
            "• Analyze top performing wells for best practices replication",
            "• Review wells with negative Net Diff BO for improvement opportunities",
            "• Monitor wells with significant performance deviations",
            "",
            "📊 Operational Actions:",
            "• Optimize production parameters for underperforming wells",
            "• Implement preventive maintenance for critical wells",
            "• Share best practices from top performers",
            "",
            "📈 Continuous Improvement:",
            "• Regular monitoring of Net Diff BO trends",
            "• Periodic review of well performance categories",
            "• Update operational strategies based on performance data"
        ]
        
        for recommendation in recommendations:
            p = text_frame.add_paragraph()
            p.text = recommendation
            p.space_after = Inches(0.03)
        
        # Save to bytes buffer
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        return ppt_buffer
        
    except Exception as e:
        st.error(f"❌ Error creating PowerPoint: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None

def create_excel_with_visualizations(data_df, stats, visualization_fig):
    """
    Create an Excel file with data, statistics, and embedded visualizations
    """
    try:
        # Create Excel writer
        excel_buffer = io.BytesIO()
        
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            # Write main data
            data_df.to_excel(writer, sheet_name='Production Data', index=False)
            
            # Write statistics
            stats_df = pd.DataFrame(list(stats.items()), columns=['Metric', 'Value'])
            stats_df.to_excel(writer, sheet_name='Statistics', index=False)
            
            # Get workbook and worksheets
            workbook = writer.book
            
            # Format worksheets
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Format data sheet
            data_sheet = writer.sheets['Production Data']
            for col_num, value in enumerate(data_df.columns.values):
                data_sheet.write(0, col_num, str(value), header_format)
            data_sheet.set_column('A:Z', 15)
            
            # Format statistics sheet
            stats_sheet = writer.sheets['Statistics']
            stats_sheet.write(0, 0, 'Metric', header_format)
            stats_sheet.write(0, 1, 'Value', header_format)
            stats_sheet.set_column('A:A', 35)
            stats_sheet.set_column('B:B', 20)
            
            # Add visualization if available
            if visualization_fig:
                # Save figure to bytes
                img_buffer = io.BytesIO()
                visualization_fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
                img_buffer.seek(0)
                
                # Create visualization sheet
                viz_sheet = workbook.add_worksheet('Visualizations')
                
                # Insert the image
                viz_sheet.insert_image('A1', 'visualization.png', {'image_data': img_buffer})
                viz_sheet.set_column('A:A', 50)
                viz_sheet.set_row(0, 300)
        
        excel_buffer.seek(0)
        return excel_buffer
        
    except Exception as e:
        st.error(f"❌ Error creating Excel file: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="Production Analysis Dashboard", 
        page_icon="🛢️", 
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS for better styling
    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .info-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin: 1rem 0;
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<h1 class="main-header">🛢️ Production Analysis Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar for navigation and info
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/179/179429.png", width=80)
        st.title("Navigation")
        st.markdown("---")
        
        st.subheader("📋 About This App")
        st.markdown("""
        This dashboard helps you analyze well production data with focus on:
        - **Net BO Production**
        - **Net Diff BO Performance**
        - **Well Performance Categories**
        - **Top Performing Wells**
        """)
        
        st.markdown("---")
        st.subheader("🎯 Quick Actions")
        if st.button("Clear Cache", help="Refresh the application"):
            st.runtime.legacy_caching.clear_cache()
            st.success("Cache cleared!")
        
        st.markdown("---")
        st.subheader("📞 Support")
        st.markdown("""
        Need help?
        - Check the instructions below
        - Ensure your Excel file format is correct
        - Contact support if issues persist
        """)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        <div class="info-box">
        <h3>🚀 Get Started</h3>
        <p>Upload your production Excel file to begin analysis. The app will automatically detect the required columns and generate comprehensive reports.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="info-box">
        <h3>📊 Expected Output</h3>
        <p>• Data Analysis<br>• Performance Visualizations<br>• Downloadable Reports<br>• Actionable Insights</p>
        </div>
        """, unsafe_allow_html=True)
    
    # File upload section
    st.markdown("---")
    st.header("📁 Upload Production Data")
    
    uploaded_file = st.file_uploader(
        "Choose your production Excel file", 
        type=['xlsx', 'xls'],
        help="Upload an Excel file with production data. The app will automatically detect the required columns."
    )
    
    if uploaded_file is not None:
        try:
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("🔍 Reading and analyzing file...")
            progress_bar.progress(25)
            
            with st.spinner("Processing your data..."):
                result_df, well_count, stats, original_columns, all_wells_data = extract_wells_with_net_diff_bo(uploaded_file)
            
            progress_bar.progress(50)
            status_text.text("📊 Generating visualizations...")
            
            if result_df is not None and not result_df.empty:
                progress_bar.progress(75)
                
                # Success message
                st.success(f"✅ Analysis complete! Processed {stats['Total All Wells']} total wells and {well_count} wells with significant Net Diff BO values.")
                
                # Main results section
                st.markdown("---")
                st.header("📈 Analysis Results")
                
                # Quick overview in columns
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Wells", stats['Total All Wells'])
                with col2:
                    st.metric("Wells with Changes", stats['Total Wells with Non-Zero Net Diff BO'])
                with col3:
                    st.metric("Positive Performance", stats['Positive Net Diff BO Wells'])
                with col4:
                    st.metric("Needs Attention", stats['Negative Net Diff BO Wells'])
                
                # Data preview
                st.subheader("📋 Production Data Preview")
                st.dataframe(result_df, use_container_width=True)
                
                # Visualizations
                st.subheader("📊 Performance Dashboard")
                data_without_total = result_df[
                    (result_df[original_columns[0]] != 'TOTAL (Non-Zero Wells)') & 
                    (result_df[original_columns[0]] != 'TOTAL (All Wells)')
                ]
                
                fig = create_visualizations(data_without_total, original_columns, all_wells_data)
                if fig:
                    st.pyplot(fig)
                else:
                    st.info("Visualizations not available due to insufficient data")
                
                progress_bar.progress(100)
                status_text.text("✅ Analysis complete!")
                
                # Download section
                st.markdown("---")
                st.header("💾 Download Reports")
                
                download_col1, download_col2 = st.columns(2)
                
                with download_col1:
                    st.subheader("📥 Data Export")
                    # CSV Download
                    csv = result_df.to_csv(index=False)
                    st.download_button(
                        label="Download CSV Report",
                        data=csv,
                        file_name="production_analysis.csv",
                        mime="text/csv",
                        help="Download the complete data analysis as CSV"
                    )
                    
                    # Excel Download
                    if st.button("Generate Excel Report", type="primary", use_container_width=True):
                        with st.spinner("Creating Excel report..."):
                            excel_buffer = create_excel_with_visualizations(result_df, stats, fig)
                        
                        if excel_buffer:
                            st.download_button(
                                label="Download Excel Report",
                                data=excel_buffer,
                                file_name="production_analysis.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Download comprehensive Excel report with data and visualizations"
                            )
                        else:
                            st.error("Failed to create Excel report")
                
                with download_col2:
                    st.subheader("📊 Presentation")
                    # PowerPoint Download
                    if st.button("Generate PowerPoint Report", type="primary", use_container_width=True):
                        with st.spinner("Creating PowerPoint presentation..."):
                            ppt_buffer = create_comprehensive_powerpoint(result_df, well_count, stats, original_columns, fig)
                        
                        if ppt_buffer:
                            st.download_button(
                                label="Download PowerPoint",
                                data=ppt_buffer,
                                file_name="production_presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                help="Download professional PowerPoint presentation"
                            )
                        else:
                            st.error("Failed to create PowerPoint presentation")
                
            else:
                st.error("❌ No valid data found in the uploaded file. Please check your file format and try again.")
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.info("💡 Tip: Make sure your Excel file has the correct format with production data in the 'Report' sheet.")
    
    else:
        # Enhanced instructions when no file is uploaded
        st.markdown("---")
        st.header("📖 How to Use This Dashboard")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🎯 Step-by-Step Guide")
            steps = [
                "1. **Prepare your Excel file** with production data",
                "2. **Upload the file** using the uploader above",
                "3. **Wait for automatic analysis** - the app detects columns automatically",
                "4. **Review the results** - data, statistics, and visualizations",
                "5. **Download reports** - CSV, Excel, or PowerPoint formats"
            ]
            
            for step in steps:
                st.markdown(step)
        
        with col2:
            st.subheader("📋 File Requirements")
            requirements = [
                "• Excel file (.xlsx or .xls)",
                "• Data in 'Report' worksheet",
                "• Multi-level headers (skip first 6 rows)",
                "• Required columns:",
                "  - Field information",
                "  - Running wells names", 
                "  - Net BO production",
                "  - Net Diff BO values"
            ]
            
            for req in requirements:
                st.markdown(req)
        
        st.markdown("---")
        st.subheader("🚀 Ready to Start?")
        st.markdown("Upload your production Excel file above to begin your analysis!")

if __name__ == "__main__":
    main()
