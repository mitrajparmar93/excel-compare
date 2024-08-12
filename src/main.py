from compare import compare_and_update_files

def main():
    DIRECTORY = 'src/data/import'
    LOCAL_FILE = 'src/data/local/my report.xlsx'
    compare_and_update_files(DIRECTORY, LOCAL_FILE)

if __name__ == "__main__":
    main()