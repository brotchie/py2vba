from py2vba import vbast
from excelbt import vbproject

def add_procedural_module_to_vbproject(project, module):
    project.add_module(vbproject.Module(module.name, module.as_code()))
    for support_module in module.support_modules:
        if isinstance(support_module, vbast.ProceduralModule):
            project.add_module(vbproject.Module(support_module.name, support_module.as_code()))
        elif isinstance(support_module, vbast.ClassModule):
            project.add_module(vbproject.ClassModule(support_module.name, support_module.as_code()))

    if module.class_support_module:
        project.add_module(vbproject.Module(module.class_support_module.name, module.class_support_module.as_code()))
    return project
