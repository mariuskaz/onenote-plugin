export default {
    template: `
        <div class="view-main">
			<h3 class="ms-font-xl"  style="margin-top: -20px;"> Connect to Todoist with personal token </h3>
			<p style="margin:40px 0;">
				<img width="30" height="30" src="../../assets/todoist-32.png" /><span class="logo">todoist</span>
			</p>
			<div class="ms-TextField">
				<label class="ms-Label">API token</label>
				<input id="token" class="ms-TextField-field" type="text" value="" placeholder="" autocomplete="on">
			</div>
			<div style="margin-top: 20px;">
				<button action="connect" class="ms-Button ms-Button--primary btn-color-red">
					<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
					<span class="ms-Button-label .btn-color-red">Next</span> 
					<span class="ms-Button-description">Connect to Todoist</span> 
				</button>
				<button action="close" class="ms-Button close">
					<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
					<span class="ms-Button-label">Cancel</span> 
				</button>
				<div class="ms-MessageBar" style="margin:40px 0">
					<div class="ms-MessageBar-content">
						<div class="ms-MessageBar-icon">
							<i class="ms-Icon ms-Icon--Info"></i>
						</div>
						<div class="ms-MessageBar-text">
							Get your token from 
							<a class="ms-Link" target="_blank" href="https://todoist.com/app/settings/integrations">Todoist integrations</a>
						</div>
					</div>
				  </div>
			</div>
		</div>
    `
}