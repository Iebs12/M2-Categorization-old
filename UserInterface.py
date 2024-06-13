import streamlit as st
import pandas as pd
import os
import requests

# Ensure the uploads directory exists to store the uploaded files
os.makedirs("uploads", exist_ok=True)

def save_uploaded_file(uploaded_file):
    """
    Save the uploaded file to the 'uploads' directory.
    
    Parameters:
        uploaded_file: The file uploaded by the user.
        
    Returns:
        str: The file path where the file is saved.
    """
    file_path = os.path.join("uploads", uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

def main():
    """
    Main function to run the Streamlit application.
    It sets up the UI for uploading an Excel file, entering a query, and 
    displays results after predicting relevancy.
    """
    # Configure the Streamlit page
    st.set_page_config(page_title="Relevancy Predictor", layout="wide")
    st.title("Relevancy Predictor")

    # Sidebar for uploading and selecting data
    st.sidebar.header("Upload and Select Data")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Save the uploaded file and get its path
        file_path = save_uploaded_file(uploaded_file)

        # Function to render the search fields
        def render_search_fields(num_fields):
            search_terms = []
            col1, col2 = st.columns([1, 3])
            for i in range(num_fields):
                with col1:
                    cat_string = st.text_input(f'Category {i+1}', key=f'cat_{i}')

                with col2:
                    query_string = st.text_input(f'Query {i+1}', key=f'query_{i}')
                search_terms.append((cat_string, query_string))
            return search_terms

        # Initialize session state for the number of search fields
        if 'num_fields' not in st.session_state:
            st.session_state.num_fields = 1

        # Render search fields based on the number of fields in session state
        search_terms = render_search_fields(st.session_state.num_fields)

        # Add and remove buttons
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            if st.button('Add Field'):
                st.session_state.num_fields += 1
        with col2:
            if st.button('Remove Field') and st.session_state.num_fields > 1:
                st.session_state.num_fields -= 1

        # Submit button
        if st.button('Submit'):
            st.write('Search Terms:', search_terms)

        
        # # Input for the query to predict relevancy
        # query = st.text_input("Enter the statement for predicting its relevancy with all the patents")

        # # Create columns for layout
        # col1, col2 = st.columns([3, 1])
        # with col1:
        #     # Button to trigger relevancy prediction
        #     predict_button = st.button("Relevancy")
        
        # if predict_button:
        #     if query:
        #         # Send the query and file path to the backend server
        #         response = requests.post("http://localhost:5000/", json={"query": query, "file_path": file_path})
        #         if response.status_code == 200:
        #             data = response.json()
        #             new_file_path = data.get('Path')
        #             filtered_file_path = data.get('FilteredPath')
        #             st.write("### Relevancy prediction completed.")
                    
        #             # Load the Excel file with updated relevancy predictions
        #             df = pd.read_excel(new_file_path)
        #             # Display the first 10 rows with selected columns
        #             st.dataframe(df[['Title', 'Relevancy predicted', 'Comments made']].head(10))
                    
        #             # Place the download buttons side by side in the right column
        #             with col2:
        #                 with st.container():
        #                     col2a, col2b = st.columns(2)
        #                     with col2a:
        #                         # Button to download the updated file
        #                         with open(new_file_path, "rb") as file:
        #                             st.download_button(
        #                                 label="Download",
        #                                 data=file,
        #                                 file_name="updated_file.xlsx",
        #                                 mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        #                             )
        #                     with col2b:
        #                         # Button to download the filtered file with relevant data only
        #                         with open(filtered_file_path, "rb") as file:
        #                             st.download_button(
        #                                 label="Download Relevant",
        #                                 data=file,
        #                                 file_name="relevant_only_file.xlsx",
        #                                 mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        #                             )
        #         else:
        #             # Display error message if the backend server returns an error
        #             st.write("Error: Could not process the query")
        #     else:
        #         # Prompt the user to enter a query if it's empty
        #         st.write("Please enter a query")

    # Custom CSS to style the buttons
    st.markdown("""
        <style>
        .stButton>button {
            background-color: #4CAF50;
            color: white;
        }
        </style>
        """, unsafe_allow_html=True)

# Run the main function when the script is executed
if __name__ == "__main__":
    main()
