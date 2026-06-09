"""
项目保存/加载系统 (PRD 5.7 桌面版简化)
=======================================
管理安全评估项目的保存、加载和版本追踪
"""

import os
import json
import shutil
from datetime import datetime


PROJECTS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'userdata', 'projects'
)


def _ensure_projects_dir():
    """Ensure projects directory exists."""
    os.makedirs(PROJECTS_DIR, exist_ok=True)


def get_project_path(name: str) -> str:
    """Get full path for a project file."""
    safe_name = name.strip().replace(' ', '_').replace('/', '_').replace('\\', '_')
    if not safe_name:
        safe_name = f'project_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    return os.path.join(PROJECTS_DIR, f'{safe_name}.json')


# ── Project Schema ──

_PROJECT_SCHEMA_VERSION = 1


def create_project(
    name: str,
    product_name: str = '',
    product_category: str = '',
    application_site: str = '',
    population: str = '成人',
    body_weight_kg: float = 60.0,
    manufacturer: str = '',
    applicant: str = '',
    market: str = 'CN',
    formula: list | None = None,
    assessment_result: dict | None = None,
    notes: str = '',
) -> dict:
    """Create a new project data structure."""
    return {
        'schema_version': _PROJECT_SCHEMA_VERSION,
        'project_name': name,
        'created_at': datetime.now().isoformat(),
        'updated_at': datetime.now().isoformat(),
        'product_info': {
            'product_name': product_name,
            'product_category': product_category,
            'application_site': application_site,
            'population': population,
            'body_weight_kg': body_weight_kg,
            'manufacturer': manufacturer,
            'applicant': applicant,
            'market': market,
        },
        'formula': formula or [],
        'assessment': assessment_result or {},
        'notes': notes,
        'history': [],
    }


def save_project(project: dict) -> str:
    """Save project to disk. Returns the file path."""
    _ensure_projects_dir()
    project['updated_at'] = datetime.now().isoformat()
    name = project.get('project_name', 'unnamed')
    path = get_project_path(name)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(project, f, ensure_ascii=False, indent=2)
    return path


def load_project(path: str) -> dict | None:
    """Load project from disk."""
    if not os.path.exists(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        project = json.load(f)
    return project


def list_projects() -> list[dict]:
    """List all saved projects."""
    _ensure_projects_dir()
    projects = []
    for fname in os.listdir(PROJECTS_DIR):
        if fname.endswith('.json'):
            path = os.path.join(PROJECTS_DIR, fname)
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    project = json.load(f)
                projects.append({
                    'path': path,
                    'name': project.get('project_name', fname),
                    'product_name': project.get('product_info', {}).get('product_name', ''),
                    'updated_at': project.get('updated_at', ''),
                    'created_at': project.get('created_at', ''),
                    'passed': project.get('assessment', {}).get('passed', 0),
                    'total': project.get('assessment', {}).get('total_ingredients', 0),
                })
            except (json.JSONDecodeError, IOError):
                continue
    # Sort by updated_at descending
    projects.sort(key=lambda p: p.get('updated_at', ''), reverse=True)
    return projects


def delete_project(path: str) -> bool:
    """Delete a project file."""
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def export_project_as_json(project: dict) -> str:
    """Export project as JSON string."""
    return json.dumps(project, ensure_ascii=False, indent=2)


def import_project_from_json(json_str: str) -> dict | None:
    """Import project from JSON string."""
    try:
        project = json.loads(json_str)
        if 'project_name' not in project:
            return None
        return project
    except json.JSONDecodeError:
        return None


def archive_project(path: str) -> str:
    """Archive (move) project to an archive subdirectory."""
    _ensure_projects_dir()
    archive_dir = os.path.join(PROJECTS_DIR, '_archive')
    os.makedirs(archive_dir, exist_ok=True)
    dest = os.path.join(archive_dir, os.path.basename(path))
    shutil.move(path, dest)
    return dest


def duplicate_project(path: str, new_name: str) -> str | None:
    """Duplicate a project with a new name."""
    project = load_project(path)
    if not project:
        return None
    project['project_name'] = new_name
    project['created_at'] = datetime.now().isoformat()
    project['updated_at'] = datetime.now().isoformat()
    return save_project(project)


def update_project_from_assessment(project: dict, assessment_result: dict) -> dict:
    """Update project with new assessment results."""
    project['assessment'] = assessment_result
    project['updated_at'] = datetime.now().isoformat()
    return project


def update_project_formula(project: dict, formula: list) -> dict:
    """Update project with formula snapshot."""
    project['formula'] = formula
    project['updated_at'] = datetime.now().isoformat()
    return project
