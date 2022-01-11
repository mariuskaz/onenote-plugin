export default {
    template: `
        <div class="view-main ms-u-slideLeftIn40">
			<h3 class="ms-font-xl"  style="margin-top:-20px;"> Add tasks to Todoist </h3>
			<div class="ms-Persona">
				<div class="ms-Persona-imageArea">
					<img id="avatar" class="ms-Persona-image" src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/images/persona.person.png">
				</div>
				<div class="ms-Persona-details">
					<div id="user" class="ms-Persona-primaryText">User name</div>
					<div id="mail" class="ms-Persona-secondaryText">mail</div>
				</div>
			</div>
			<div id="dropdown" class="ms-Dropdown" tabindex="0" style="margin-top:30px;">
				<i class="ms-Dropdown-caretDown ms-Icon ms-Icon--ChevronDown"></i>
				<select change="refresh" id="projects" class="ms-Dropdown-select">
					<option value="new">Create new project</option>
				</select>
			</div>
			<div class="ms-MessageBar ms-MessageBar--success" style="margin-top: 10px;">
				<div class="ms-MessageBar-content">
					<div class="ms-MessageBar-icon">
						<i class="ms-Icon ms-Icon--Completed"></i>
					</div>
					<div class="ms-MessageBar-text">
						<b id="tasks">Tasks</b> found on this page and will be added to Todoist.
						<a class="ms-Link" href="#" click="retry">Refresh</a>
					</div>
				</div>
			</div>
			<div class="ms-Toggle" style="margin:20px 0px 30px;">
				<span class="ms-Toggle-description">Link tasks to this page</span> 
				<input type="checkbox" id="toggle-links" class="ms-Toggle-input" />
				<label for="toggle-links" id="add-links" click="toggle" class="ms-Toggle-field" tabindex="0">
					<span class="ms-Label ms-Label--off">Off</span> 
					<span class="ms-Label ms-Label--on">On</span> 
				</label>
			</div>
			<h3 class="ms-font-xl"> Preview </h3>
			<div>
				<h3 id="title">Inbox</h3>
				<hr style="border:none;background-color:lightgray;height:1px;margin-top:-5px">
				<p style="padding:5px;margin-top:15px;">&#9898;&nbsp;&nbsp;&nbsp;<span id="preview"></span></p>
			</div>
			<div style="margin-top: 40px;">
				<button click="push" class="ms-Button ms-Button--primary">
					<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
					<span class="ms-Button-label">Export</span> 
				</button>
				<button click="close" class="ms-Button">
					<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
					<span class="ms-Button-label">Cancel</span> 
				</button>
			</div>
			<button click="disconnect" class="ms-Button ms-Button--compound" style="margin-top:50px">
				<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
				<span class="ms-Button-label">Disconnect</span> 
				<span class="ms-Button-description">from current connected Todoist account</span> 
			</button>
		</div>
    `
}