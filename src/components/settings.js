export default {
    template: `
        <div class="view-main">
			<h3 class="ms-font-xl"  style="margin-top:-20px;"> Add tasks to Todoist </h3>
			<div class="ms-Persona">
				<div class="ms-Persona-imageArea">
					<img id="avatar" class="ms-Persona-image" src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/images/persona.person.png">
				</div>
				<div class="ms-Persona-presence">
					<i class="ms-Persona-presenceIcon ms-Icon ms-Icon--SkypeCheck"></i>
				</div>
				<div class="ms-Persona-details">
					<div id="user" class="ms-Persona-primaryText">User name</div>
					<div id="mail" class="ms-Persona-secondaryText">mail</div>
				</div>
			</div>
			<div id="dropdown" class="ms-Dropdown" tabindex="0" style="margin-top: 30px;">
				<i class="ms-Dropdown-caretDown ms-Icon ms-Icon--ChevronDown"></i>
				<select id="projects" class="ms-Dropdown-select ms-TextField--underlined">
					<option value="none">Create new project</option>
				</select>
			</div>
			<div class="ms-MessageBar ms-MessageBar--success" style="margin-top: 10px;">
				<div class="ms-MessageBar-content">
					<div class="ms-MessageBar-icon">
						<i class="ms-Icon ms-Icon--Completed"></i>
					</div>
					<div class="ms-MessageBar-text">
						<b id="tasks">Tasks</b> found on this page and will be added to Todoist project.
					</div>
				</div>
			</div>
			<div class="ms-Toggle" style="margin-top: 20px;">
				<span class="ms-Toggle-description">Link tasks to this page</span> 
				<input type="checkbox" id="demo-toggle-1" class="ms-Toggle-input" />
				<label for="demo-toggle-1" class="ms-Toggle-field" tabindex="0">
					<span class="ms-Label ms-Label--off">Off</span> 
					<span class="ms-Label ms-Label--on">On</span> 
				</label>
			</div>
			<div style="margin-top: 40px;">
				<button action="push" class="ms-Button ms-Button--primary">
					<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
					<span class="ms-Button-label">Export</span> 
				</button>
				<button action="close" class="ms-Button">
					<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
					<span class="ms-Button-label">Cancel</span> 
				</button>
			</div>
		</div>
    `
}