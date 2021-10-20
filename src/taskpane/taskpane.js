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
	avatar: "../../assets/avatar.png",
	pageTitle: "Onenote",
	pageUrl: location.href,
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

	toggle() {
		view.addLinks = !view.addLinks
	},

	push() {
		let projects = view.get('projects')
		todoist.push(projects.value, view.tasks)
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
					avatar: data.user.avatar_medium || view.avatar,
					user: data.user.full_name,
					mail: data.user.email,
					tasks: view.tasks.length + " task(s)",
					project: view.pageTitle,
					task: view.tasks[0].substring(0, 34),
				})

				let list = view.get("projects")
				data.projects.forEach(project => {
					list.options.add(new Option(project.name, project.id))
				})

				let projects = view.get("dropdown")
				new fabric['Dropdown'](projects)

				if (view.addLinks) {
					let links = view.get("add-links")
					if (links != null) links.classList.add("is-selected")
				}
			})
		})

		.catch(error => {
			view.alert("Connection failed!", error);
		})
	},

	push(project = "new", tasks = []) {
		
		let item = 0, headers = {
			'Authorization': 'Bearer ' + this.token,
			'Content-Type': 'application/json'
		}

		view.show(status)
		if (project == "new") {
			fetch('https://api.todoist.com/rest/v1/projects', { 
				method: 'POST',
				headers : headers,
				body: JSON.stringify({ name: view.pageTitle })
			})
			.then(res => res.json())
			.then(project => todoist.push(project.id, tasks))

		} else {

			let projectId = parseInt(project)
			console.log('projectId:', projectId)

			tasks.forEach( task => {
				let todo = view.addLinks ? 
					`[${task}](${view.pageUrl})` : task

				fetch('https://api.todoist.com/rest/v1/tasks', { 
					method: 'POST',
					headers : headers,
					body: JSON.stringify({ content: todo, project_id: projectId })
				})

				.then(res => {
					item ++
					console.log('tasks pushed:', res.ok)
					if (item == tasks.length) {
						window.open("https://todoist.com/showProject?id=" + projectId, "_blank")
						view.close()
					}
				})

				.catch(error => {
					view.alert("Task"+ item +" push failed!", error);
				})
			})
			
		}
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
		outlines = [],
		tables = [],
		strings = []

		page.load("title, webUrl")
		page.contents.load("type, items");
		return context.sync()

		.then(() => {
			view.pageTitle = page.title
			view.pageUrl = page.webUrl
			page.contents.items.forEach(item => {
				if (item.type == 'Outline') {
					item.load("outline/paragraphs, outline/paragraphs/type")
					outlines.push(item)
				}
			})
			return context.sync()
		})

		.then(() => {
			outlines.forEach(item => {
				item.outline.paragraphs.items.forEach( p => {
					if (p.type == "RichText") {
						p.load("richtext")
						strings.push(p.richText.getHtml())
					}

					/* https://docs.microsoft.com/en-us/javascript/api/onenote/onenote.table?view=onenote-js-1.1 */

					if (p.type == "Table") {
						p.load("table/rows/items/cells/paragraphs/type") 
						tables.push(p.table)
					}
				})
			})
			return context.sync()
		})

		.then(() => {
			tables.forEach( table => {
				table.rows.items.forEach( row => {
					row.cells.items.forEach( cell => {
						cell.paragraphs.items.forEach( p => {
							if (p.type == "RichText") {
								p.load("richtext")
								strings.push(p.richText.getHtml())
							}
						})
					})
				}) 
			})
			return context.sync()
		})

		.then(() => {
			strings.forEach( html => {
				let doc = parser.parseFromString(html.value, 'text/html'),
				tag = doc.querySelector("[data-tag=to-do]")
				if (tag != null) view.tasks.push(tag.innerText)
			})

			console.log('tasks found:', view.tasks.length)
			//todoist.token = "none"

			if (view.tasks.length == 0)
				view.alert("No tasks found! There is nothing to export.", 
						   "No to-do tags found on this page!")
			else if (todoist.token == "none") view.show(connect)
			else todoist.sync()
			
		})
		
	})

	.catch(error => {
		view.alert("Error!", error);
	})

}
