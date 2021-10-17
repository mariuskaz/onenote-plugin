/* eslint-disable */
import "../../assets/todoist-16.png"
import "../../assets/todoist-32.png"
import "../../assets/todoist-80.png"
import "../../assets/gantt-16.png"
import "../../assets/gantt-32.png"
import "../../assets/gantt-80.png"

import alert from "../components/alert.js"
import status from "../components/status.js"
import connect from "../components/connect.js"
import settings from "../components/settings.js"

let view = {

	addLinks: false,
	actions: ["click", "change"],
	pageTitle: "Onenote",
	tasks: [],

	get(id) {
		return document.getElementById(id)
	},

	update(data) {
		for (let id in data) {
			let el = this.get(id)
			switch (el.nodeName) {
				case 'INPUT':
					el.value = data[id]
					break
				case 'IMG':
					el.src = data[id]
					break
				default:
					el.innerHTML = data[id]
			}
		}
   	},
 
	show(component) {
		this.get("app-body").innerHTML = component.template

		view.actions.forEach( action => {
			let elements = this.get('app-body').querySelectorAll("["+action+"]")
			elements.forEach( el => {
				el.addEventListener(action, view[el.getAttribute(action)])
			})
		})

		let toggles = document.querySelectorAll(".ms-Toggle")
		toggles.forEach( toggle => new fabric['Toggle'](toggle) )
		return this
	},

	alert(title, details = "unknown") {
		this.show(alert).update({ title, details })
	},

	connect() {
		let token = view.get("token").value
		if (token.length > 0) todoist.sync(token)
	},

	push() {
		view.show(status)
		todoist.push(view.tasks)
	},

	retry() {
		getPageTasks()
	},

	refresh() {
		let projects = view.get("projects"),
		project = projects.selectedOptions[0].text
		if (projects.value == "new") project = view.pageTitle
		view.update({ project })
	},

	close() {
		Office.context.ui.closeContainer()
	},

},

todoist = {

	token: localStorage["todoist_token"] || "none",

	sync(token) {
		if (token) this.token = token

		let sync_url = "https://api.todoist.com/sync/v8/sync",
		headers = {
			'Authorization': 'Bearer ' + this.token,
			'Content-Type': 'application/json',
		},

		params = {
			sync_token: '*',
			resource_types: ["all"]
		}

		view.show(status)
		fetch(sync_url, { 
			method: 'POST',
			headers : headers,
			body: JSON.stringify(params)
		})
	
		.then(res => {
			res.json().then(data => {
				localStorage.setItem("todoist_token", todoist.token)

				view.show(settings).update({
					avatar: data.user.avatar_medium,
					user: data.user.full_name,
					mail: data.user.email,
					tasks: view.tasks.length + " task(s)",
					project: view.pageTitle,
					task: view.tasks[0],
				})

				let list = view.get("projects")
				data.projects.forEach(project => {
					list.options.add(new Option(project.name, project.id))
				})

				let projects = view.get("dropdown")
				new fabric['Dropdown'](projects)
			})
		})

		.catch(error => {
			view.alert("Connection failed!", error);
		})
	},

	push(tasks = []) {
		
		let item = 0,
		headers = {
			'Authorization': 'Bearer ' + this.token,
			'Content-Type': 'application/json'
		}

		tasks.forEach( todo => {
			fetch('https://api.todoist.com/rest/v1/tasks', { 
				method: 'POST',
				headers : headers,
				body: JSON.stringify({ content: todo })
			})

			.then(res => {
				item ++
				console.log("status: ", res.status)
				if (item == tasks.length) {
					view.close()
					window.open("https://todoist.com", "_blank")
				}
			})

			.catch(error => {
				view.alert("Task"+ item +" push failed!", error);
			})

		})
	},

}

Office.onReady((info) => {
	if (info.host === Office.HostType.OneNote) getPageTasks()
})

export async function getPageTasks() {

	view.tasks = []
	view.show(status)
	OneNote.run(context => {

		let parser = new DOMParser(),
		page = context.application.getActivePage(),
		outlines = []

		page.load("title")
		page.contents.load("items");
		return context.sync().then(() => {
			view.pageTitle = page.title
			console.log("checking outlines...");
			page.contents.items.forEach(item => {
				outlines.push(item)
				item.outline.paragraphs.load("items");
			})

			return context.sync().then(() => {
				let strings = [], tables = []
				console.log("checking paragraphs...");
				outlines.forEach( item => {
					item.outline.paragraphs.items.forEach( p => {
						if (p.type == "RichText"){
							let html = p.richText.getHtml();
							strings.push(html)
							p.load("richtext");
						}
						if (p.type == "Table"){
							tables.push(p.table)
							p.load("table");
							let c = p.table.getCell(1,1)
							//console.log("cell", c)
						}
					})
				})

				return context.sync().then(function(){
					strings.forEach( html => {
						let doc = parser.parseFromString(html.value, 'text/html'),
						tag = doc.querySelector("[data-tag=to-do]")
						if (tag != null) view.tasks.push(tag.innerText)
					})

					/* https://docs.microsoft.com/en-us/javascript/api/onenote/onenote.table?view=onenote-js-1.1 */
					tables.forEach( table => {
						console.log('table rows:', table.rowCount)
						let c = table.getCell(1, 1)
						//console.log(c.paragraphs.items.length)
					})

					console.log('tasks found:', view.tasks.length)
					todoist.token = "none"

					if (view.tasks.length == 0) {
						view.alert("No tasks found! There is nothing to export.", 
								   "No to-do tags found on this page!")

					} else if (todoist.token == "none") {
						view.show(connect)

					} else {
						todoist.sync()
					}
					
				})
			})
		})

		.catch(error => {
			view.alert("Error!", error);
		})

	})
}