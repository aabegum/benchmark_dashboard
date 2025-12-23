import os

def create_structure(base_dir='dashboard'):
    # Create base directory
    os.makedirs(base_dir, exist_ok=True)
    
    # Create main.py
    open(os.path.join(base_dir, 'main.py'), 'w').close()
    
    # Create config directory and files
    config_dir = os.path.join(base_dir, 'config')
    os.makedirs(config_dir, exist_ok=True)
    open(os.path.join(config_dir, '__init__.py'), 'w').close()
    open(os.path.join(config_dir, 'settings.py'), 'w').close()
    open(os.path.join(config_dir, 'config.yaml'), 'w').close()
    
    # Create data directory and files
    data_dir = os.path.join(base_dir, 'data')
    os.makedirs(data_dir, exist_ok=True)
    open(os.path.join(data_dir, '__init__.py'), 'w').close()
    open(os.path.join(data_dir, 'loader.py'), 'w').close()
    open(os.path.join(data_dir, 'processor.py'), 'w').close()
    open(os.path.join(data_dir, 'validators.py'), 'w').close()
    
    # Create analysis directory and files
    analysis_dir = os.path.join(base_dir, 'analysis')
    os.makedirs(analysis_dir, exist_ok=True)
    open(os.path.join(analysis_dir, '__init__.py'), 'w').close()
    open(os.path.join(analysis_dir, 'outliers.py'), 'w').close()
    open(os.path.join(analysis_dir, 'statistics.py'), 'w').close()
    open(os.path.join(analysis_dir, 'filters.py'), 'w').close()
    
    # Create visualization directory and files
    visualization_dir = os.path.join(base_dir, 'visualization')
    os.makedirs(visualization_dir, exist_ok=True)
    open(os.path.join(visualization_dir, '__init__.py'), 'w').close()
    open(os.path.join(visualization_dir, 'plots.py'), 'w').close()
    open(os.path.join(visualization_dir, 'charts.py'), 'w').close()
    open(os.path.join(visualization_dir, 'utils.py'), 'w').close()
    
    # Create ui directory and files
    ui_dir = os.path.join(base_dir, 'ui')
    os.makedirs(ui_dir, exist_ok=True)
    open(os.path.join(ui_dir, '__init__.py'), 'w').close()
    open(os.path.join(ui_dir, 'sidebar.py'), 'w').close()
    open(os.path.join(ui_dir, 'main_content.py'), 'w').close()
    open(os.path.join(ui_dir, 'components.py'), 'w').close()
    
    # Create utils directory and files
    utils_dir = os.path.join(base_dir, 'utils')
    os.makedirs(utils_dir, exist_ok=True)
    open(os.path.join(utils_dir, '__init__.py'), 'w').close()
    open(os.path.join(utils_dir, 'helpers.py'), 'w').close()
    open(os.path.join(utils_dir, 'constants.py'), 'w').close()
    
    # Create requirements.txt
    open(os.path.join(base_dir, 'requirements.txt'), 'w').close()

if __name__ == '__main__':
    create_structure()