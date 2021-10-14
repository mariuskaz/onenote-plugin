export default {
    template: `
        <div class="view-main">
			<h3 id="title" class="ms-font-xl" style="margin-top: -20px;"></h3>
			<div class="ms-MessageBar ms-MessageBar--error">
			<div class="ms-MessageBar-content">
				<div class="ms-MessageBar-icon">
					<i class="ms-Icon ms-Icon--ErrorBadge" style="padding:5px"></i>
				</div>
				<div id="details" class="ms-MessageBar-text" style="padding:5px 5px 5px 0"></div>
			</div>
			</div>
			<div style="margin:30px 0;">
			<button action="close" class="ms-Button ms-Button--primary close">
				<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span> 
				<span class="ms-Button-label">OK</span> 
			</button>
			</div>
		</div>
    `
}