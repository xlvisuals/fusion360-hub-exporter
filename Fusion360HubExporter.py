#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Export all Fusion 360 projects and designs from the active hub to disk
# - Exports components as .f3d, .stp, .stl, and .igs
# - Exports a screenshot of the project (optional)
# - Exports component bodies as .stl (optional)
# - Exports component sketches as .dxf (optional)

# How to use:
# 0. Configure export options in this script (below)
# 1. Open Fusion 360
# 2. Click on "Utilities" > "Add-ins" > "Scripts and Add-ins"
# 3. Click the + button and select the folder with this script
# 4. Confirm popup that this might take a while
# 5. Select target folder
# 6. Wait while fusion opens every project and design, exports it, then closes it.

# Acknowledgements:
# This program is based on the work of Justin Nesselrotte - with gratitude.
# https://github.com/Jnesselr/fusion-360-total-exporter

__author__ = "Xlvisuals Limited"
__license__ = "MIT"
__date__ = "2025-10-22"

# Options:
# Only set SKIP_PROJECT_NAMES or EXPORT_PROJECT_NAMES, not both.
SKIP_PROJECT_NAMES = []  # If set, projects with names in this list are not exported.
EXPORT_PROJECT_NAMES = []  # If set, only projects with names in this list are exports. Leave empty to export all. Ignored if SKIP_PROJECT_NAMES is set.
EXPORT_COMPONENT_FORMATS = ['stp', 'stl', 'igs']
EXPORT_SCREENSHOT = True
EXPORT_SKETCHES = True
EXPORT_BODIES = False  # Export each component body
EXPORT_SUBCOMPONENTS = True
MAX_SUBCOMPONENT_COUNT = 300  # more than this and we don't export subcomponents at all.
OVERWRITE_EXISTING = True
LOGGER_NAME = "Fusion360 Hub Exporter"

# Imports
from logging import Logger, FileHandler, Formatter
from datetime import datetime
import os
import sys
import re

try:
    # Wrap Fusion SDK import in try/except so we can log in case it fails.
    import adsk.core, adsk.fusion, adsk.cam, traceback


    class Fusion360HubExporter(object):

        def __init__(self, app):
            self.app = app
            self.active_hub_name = app.data.activeHub.name  # can only open designs from active hub
            self.ui = self.app.userInterface
            self.data = self.app.data
            self.documents = self.app.documents
            self.log = Logger(LOGGER_NAME)
            self.num_issues = 0
            self.was_cancelled = False
            self.export_projects = EXPORT_PROJECT_NAMES
            self.skip_projects = SKIP_PROJECT_NAMES
            self.export_screenshot = EXPORT_SCREENSHOT
            self.export_formats = ['stp', 'stl', 'igs']  # supported: ['stp', 'stl', 'igs']
            self.export_bodies = EXPORT_BODIES
            self.export_sketches = EXPORT_SKETCHES
            self.export_subcomponents = EXPORT_SUBCOMPONENTS
            self.max_subcomponent_count = MAX_SUBCOMPONENT_COUNT
            self.overwrite_existing = OVERWRITE_EXISTING
            self.progress_dialog = None

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            pass

        def run(self, context):
            self.ui.messageBox(
                "This plugin will block Fusion while it opens and close every project for export.\n\n" \
                "Exporting your designs will take a while, depending on how many you have.\n\n" \
                "You can cancel at any time.\n\n" \
                )

            output_path = self._ask_for_output_path()

            if output_path is None:
                return

            now = datetime.now().strftime('%Y-%m-%d')
            file_handler = FileHandler(os.path.join(output_path, 'Fusion360HubExporter_{}.log'.format(now)))
            file_handler.setFormatter(Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            self.log.addHandler(file_handler)

            self.log.info("Starting export.")
            self.log.info("Python version: {}".format(sys.version))

            self._export_data(output_path)

            self.log.info("Done exporting!")

            if self.was_cancelled:
                self.ui.messageBox("Cancelled!")
            elif self.num_issues > 0:
                self.ui.messageBox(
                    "The exporting process ran into {num_issues} issue{english_plurals}. Please check the log for more information".format(
                        num_issues=self.num_issues,
                        english_plurals="s" if self.num_issues > 1 else ""
                    ))
            else:
                self.ui.messageBox("Export finished successfully.")

        def _export_data(self, output_path):
            self.progress_dialog = self.ui.createProgressDialog()
            self.progress_dialog.show("Exporting data!", "", 0, 1, 1)

            all_hubs = self.data.dataHubs
            for hub_index in range(all_hubs.count):
                hub = all_hubs.item(hub_index)

                if hub.name == self.active_hub_name:
                    self.log.info("Exporting hub \"{}\"".format(hub.name))

                    all_projects = hub.dataProjects
                    for project_index in range(all_projects.count):
                        if self.was_cancelled or self.progress_dialog.wasCancelled:
                            self.log.info("The process was cancelled!")
                            self.was_cancelled = True
                            return

                        files = []
                        project = all_projects.item(project_index)

                        if self.skip_projects and project.name in self.skip_projects:
                            self.log.info("Skipping project \"{}\" as project in exclude list.".format(project.name))
                        elif self.export_projects and project.name not in self.export_projects:
                            self.log.info("Not exporting project \"{}\" as not in include list.".format(project.name))
                        else:
                            self.log.info("Exporting project \"{}\"".format(project.name))

                            folder = project.rootFolder

                            files.extend(self._get_files_for(folder))

                            self.progress_dialog.message = "Hub: {} of {}\nProject: {} of {}\nExporting design %v of %m".format(
                                hub_index + 1,
                                all_hubs.count,
                                project_index + 1,
                                all_projects.count
                            )
                            self.progress_dialog.maximumValue = len(files)
                            self.progress_dialog.reset()

                            if not files:
                                self.log.info("No files to export for this project")
                                continue

                            for file_index in range(len(files)):
                                if self.was_cancelled or self.progress_dialog.wasCancelled:
                                    self.log.info("The process was cancelled!")
                                    self.was_cancelled = True
                                    return

                                file: adsk.core.DataFile = files[file_index]
                                self.progress_dialog.progressValue = file_index + 1
                                self._export_design(output_path, file)
                            self.log.info("Finished exporting project \"{}\".\n\n".format(project.name))
                    self.log.info("Finished exporting hub \"{}\"\n".format(hub.name))
                else:
                    self.log.info(
                        "Skipping inactive hub \"{}\" - Fusion can only open documents from the active hub.".format(
                            hub.name)
                    )

        def _ask_for_output_path(self):
            folder_dialog = self.ui.createFolderDialog()
            folder_dialog.title = "Where do you want to save the exported projects?"
            dialog_result = folder_dialog.showDialog()
            if dialog_result != adsk.core.DialogResults.DialogOK:
                return None

            output_path = folder_dialog.folder

            return output_path

        def _get_files_for(self, folder):
            files = []
            try:
                for file in folder.dataFiles:
                    files.append(file)

                for sub_folder in folder.dataFolders:
                    files.extend(self._get_files_for(sub_folder))
            except Exception as e:
                self.log.error("Exception getting files: {}".format(e))

            return files

        def _export_design(self, root_folder, file: adsk.core.DataFile):
            if file.fileExtension != "f3d" and file.fileExtension != "f3z":
                self.log.info("Not exporting \"{}\" (file is not a Fusion Design)".format(file.name))
                return

            self.log.info("Exporting design \"{}\"".format(file.name))

            document = None
            try:
                document = self.documents.open(file)
                if document is None:
                    raise Exception("Documents.open returned None")
                document.activate()
            except BaseException as ex:
                self.num_issues += 1
                self.log.exception("Opening {} failed!".format(file.name), exc_info=ex)
                try:
                    if document is not None:
                        document.close(False)
                except BaseException as ex:
                    self.num_issues += 1
                    self.log.exception("Failed to close \"{}\"".format(file.name), exc_info=ex)
                return

            try:
                file_folder = file.parentFolder
                file_folder_path = self._cleanup_name(file_folder.name)

                while file_folder.parentFolder is not None:
                    file_folder = file_folder.parentFolder
                    file_folder_path = os.path.join(self._cleanup_name(file_folder.name), file_folder_path)

                parent_project = file_folder.parentProject
                parent_hub = parent_project.parentHub

                file_folder_path = str(self._create_path(
                    root_folder,
                    "Hub {}".format(self._cleanup_name(parent_hub.name)),
                    "Project {}".format(self._cleanup_name(parent_project.name)),
                    file_folder_path,
                    self._cleanup_name(file.name) + "." + file.fileExtension
                ))

                if not os.path.exists(file_folder_path):
                    self.num_issues += 1
                    self.log.exception("Couldn't make root folder\"{}\"".format(file_folder_path))
                    return

                self.log.info("Writing to \"{}\"".format(file_folder_path))
                file_export_path = os.path.join(file_folder_path, self._cleanup_name(file.name))
                file_screenshot_path = os.path.join(file_folder_path, self._cleanup_name(file.name) + ".png")

                if self.export_screenshot:
                    if self.overwrite_existing or not os.path.exists(file_screenshot_path):
                        try:
                            # write screenshot
                            self.app.activeViewport.refresh()
                            adsk.doEvents()
                            self.app.activeViewport.saveAsImageFile(file_screenshot_path, 1024, 1024)
                            self.log.info("Screenshot saved to \"{}\"".format(file_screenshot_path))
                        except Exception as e:
                            self.log.warning("Error saving screenshot to {}: {}".format(file_screenshot_path, e))
                    else:
                        self.log.info("Screenshot file \"{}\" already exists.".format(file_screenshot_path))

                try:
                    fusion_document: adsk.fusion.FusionDocument = adsk.fusion.FusionDocument.cast(document)
                    design: adsk.fusion.Design = fusion_document.design
                    export_manager: adsk.fusion.ExportManager = design.exportManager

                    # Write f3d/f3z file
                    options = export_manager.createFusionArchiveExportOptions(file_export_path)
                    export_manager.execute(options)

                    # Write components
                    try:
                        self._write_component(file_folder_path, design.rootComponent)
                    except Exception as e:
                        self.num_issues += 1
                        self.log.error("Error saving components to {}: {}".format(file_folder_path, e))

                except Exception as e:
                    self.num_issues += 1
                    self.log.error("Error saving design to {}: {}".format(file_export_path, e))

                self.log.info("Finished exporting file \"{}\"".format(file.name))
            except Exception as ex:
                self.num_issues += 1
                self.log.error("Error working on {}: {}".format(file.name, ex))

            finally:
                try:
                    if document is not None:
                        document.close(False)
                except BaseException as ex:
                    self.num_issues += 1
                    self.log.exception("Failed to close \"{}\"".format(file.name), exc_info=ex)

        def _write_component(self, component_base_path, component: adsk.fusion.Component):
            if self.progress_dialog and self.progress_dialog.wasCancelled:
                self.was_cancelled = True
                return

            self.log.info("Writing component \"{}\" to \"{}\"".format(component.name, component_base_path))
            design = component.parentDesign

            output_path = os.path.join(component_base_path, self._cleanup_name(component.name))

            # Export Component
            if 'stp' in self.export_formats:
                try:
                    self._write_step(output_path, component)
                except Exception as e:
                    self.log.error("Error saving STP for {}: {}".format(component.name, e))
            if 'stl' in self.export_formats:
                try:
                    self._write_stl(output_path, component)
                except Exception as e:
                    self.log.error("Error saving STL for {}: {}".format(component.name, e))
            if 'igs' in self.export_formats:
                try:
                    self._write_iges(output_path, component)
                except Exception as e:
                    self.log.error("Error saving IGS for {}: {}".format(component.name, e))

            # Export Sketches
            if self.export_sketches:
                sketches = component.sketches
                for sketch_index in range(sketches.count):
                    sketch_path = ''
                    try:
                        sketch = sketches.item(sketch_index)
                        sketch_path = os.path.join(output_path, sketch.name)
                        self._write_dxf(sketch_path, sketch)
                    except Exception as e:
                        if sketch_path:
                            self.log.error("Error saving sketch to {}: {}".format(sketch_path, e))
                        else:
                            self.log.error("Error getting sketch from {}: {}".format(component.name, e))

            # Export Bodies
            if self.export_bodies:
                bRepBodies = component.bRepBodies
                meshBodies = component.meshBodies

                if bRepBodies.count > 0:
                    self._create_path(output_path)
                    for index in range(bRepBodies.count):
                        if self.progress_dialog and self.progress_dialog.wasCancelled:
                            self.was_cancelled = True
                            return
                        body = bRepBodies.item(index)
                        self._write_stl_body(os.path.join(output_path, body.name), body)

                if meshBodies.count > 0:
                    for index in range(meshBodies.count):
                        if self.progress_dialog and self.progress_dialog.wasCancelled:
                            self.was_cancelled = True
                            return
                        body = meshBodies.item(index)
                        self._write_stl_body(os.path.join(output_path, body.name), body)

            # Export Subcomponent recursively
            if self.export_subcomponents:
                occurrences = component.occurrences
                subcomponent_count = occurrences.count
                if subcomponent_count > self.max_subcomponent_count:
                    # This is to prevent getting stuck for hours on some PCB design with thousands of components.
                    self.log.info(
                        "Component {} has {} subcomponents, which exceeds the set limit of {}. No subcomponents exported.".format(
                            component.name, subcomponent_count, self.max_subcomponent_count))
                else:
                    for occurrence_index in range(occurrences.count):
                        if self.progress_dialog and self.progress_dialog.wasCancelled:
                            self.was_cancelled = True
                            return
                        sub_component_name = ''
                        try:
                            occurrence = occurrences.item(occurrence_index)
                            sub_component = occurrence.component
                            sub_component_name = sub_component.name
                            sub_path = self._create_path(component_base_path, self._cleanup_name(component.name))

                            self._write_component(sub_path, sub_component)
                        except Exception as e:
                            if sub_component_name:
                                self.log.error("Error saving sub-component {}: {}".format(sub_component_name, e))
                            else:
                                self.log.error("Error getting sub-component from {}: {}".format(component.name, e))

        def _write_step(self, output_path, component: adsk.fusion.Component):
            file_path = output_path + ".stp"
            if not self.overwrite_existing and os.path.exists(file_path):
                self.log.info("Step file \"{}\" already exists".format(file_path))
                return

            self.log.info("Writing step file \"{}\"".format(file_path))
            export_manager = component.parentDesign.exportManager

            options = export_manager.createSTEPExportOptions(output_path, component)
            export_manager.execute(options)

        def _write_stl(self, output_path, component: adsk.fusion.Component):
            file_path = output_path + ".stl"
            if not self.overwrite_existing and os.path.exists(file_path):
                self.log.info("Stl file \"{}\" already exists".format(file_path))
                return

            self.log.info("Writing stl file \"{}\"".format(file_path))
            export_manager = component.parentDesign.exportManager

            try:
                options = export_manager.createSTLExportOptions(component, output_path)
                export_manager.execute(options)
            except BaseException as ex:
                self.log.exception("Failed writing stl file \"{}\"".format(file_path), exc_info=ex)

                if component.occurrences.count + component.bRepBodies.count + component.meshBodies.count > 0:
                    self.num_issues += 1

        def _write_stl_body(self, output_path, body):
            file_path = output_path + ".stl"
            if not self.overwrite_existing and os.path.exists(file_path):
                self.log.info("Stl body file \"{}\" already exists".format(file_path))
                return

            self.log.info("Writing stl body file \"{}\"".format(file_path))
            export_manager = body.parentComponent.parentDesign.exportManager

            try:
                options = export_manager.createSTLExportOptions(body, file_path)
                export_manager.execute(options)
            except BaseException:
                # Probably an empty model, ignore it
                pass

        def _write_iges(self, output_path, component: adsk.fusion.Component):
            file_path = output_path + ".igs"
            if not self.overwrite_existing and os.path.exists(file_path):
                self.log.info("Iges file \"{}\" already exists".format(file_path))
                return

            self.log.info("Writing iges file \"{}\"".format(file_path))

            export_manager = component.parentDesign.exportManager

            options = export_manager.createIGESExportOptions(file_path, component)
            export_manager.execute(options)

        def _write_dxf(self, output_path, sketch: adsk.fusion.Sketch):
            file_path = output_path + ".dxf"
            if not self.overwrite_existing and os.path.exists(file_path):
                self.log.info("DXF sketch file \"{}\" already exists".format(file_path))
                return

            self.log.info("Writing dxf sketch file \"{}\"".format(file_path))

            sketch.saveAsDXF(file_path)

        def _create_path(self, *path):
            out_path = os.path.join(*path)
            os.makedirs(out_path, exist_ok=True)
            return out_path

        def _cleanup_name(self, name):
            name = re.sub(r'[^a-zA-Z0-9 \n\.]', ' ', name).strip()

            if name.endswith('.stp') or name.endswith('.stl') or name.endswith('.igs'):
                name = name[0: -4] + "_" + name[-3:]

            return name


    def run(context):
        ui = None
        try:
            app = adsk.core.Application.get()

            with Fusion360HubExporter(app) as total_export:
                total_export.run(context)

        except:
            ui = app.userInterface
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


except Exception as e:
    file_folder = os.path.dirname((os.path.realpath(__file__)))
    file_handler = FileHandler(os.path.join(file_folder, 'error.log'))
    file_handler.setFormatter(Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    errorlog = Logger(LOGGER_NAME)
    errorlog.addHandler(file_handler)
    errorlog.error("Error in Fusion360HubExporter: {}".format(e))
    exit(1)
