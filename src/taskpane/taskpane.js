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

	getValue(id) {
		let el = this.get(id)
		return el.value || el.innerText || ""
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
					el.innerText = data[id]
			}
		}
	},
 
	show(component, data) {
		this.get("app-body").innerHTML = component.template

		view.actions.forEach( action => {
			let elements = this.get('app-body').querySelectorAll(`[${action}]`)
			elements.forEach( el => {
				el.addEventListener(action, view[el.getAttribute(action)])
			})
		})

		let toggles = document.querySelectorAll(".ms-Toggle")
		toggles.forEach( toggle => new fabric['Toggle'](toggle) )

		if (data != undefined) this.update(data)
		return this
	},

	alert(title, details = "unknown") {
		this.show(alert, { title, details })
	},

	connect() {
		let token = view.getValue("token")
		if (token.length > 0) todoist.sync(token)
	},

	disconnect() {
		localStorage.removeItem('todoist_token')
		todoist.token = 'none'
		view.close()
	},

	toggle() {
		view.addLinks = !view.addLinks
	},

	push() {
		let projectId = view.getValue('projects')

		if (projectId == "new") {
			todoist.push(view.tasks)

		} else {
			todoist.getData(projectId).then(data => {
				let items = [], todos = []
				data.items.forEach( item => items.push(tools.getText(item.content)) )
				todos = view.tasks.filter( todo => !items.includes(todo) )
				todoist.push(todos, projectId)
			})

		}

	},

	retry() {
		getPageTasks()
	},

	refresh() {
		let projects = view.get("projects"),
		title = projects.selectedOptions[0].text,
		tasks = view.tasks.length + " new task(s)"

		if (projects.value == "new") {
			title = view.pageTitle
			view.update({ title, tasks	})
		
		} else {
			todoist.getData(projects.value).then(data => {
				let items = [], todos = []
				data.items.forEach( item => items.push(tools.getText(item.content)) )
				todos = view.tasks.filter( todo => !items.includes(todo) )
				tasks = todos.length + " new task(s)"
				view.update({ title, tasks })
			})
		}
			
	},

	close() {
		Office.context.ui.closeContainer()
	},

},

todoist = {

	token: localStorage["todoist_token"] || "none",
	url: "https://api.todoist.com/sync/v8/sync",

	sync(token) {
		if (token) this.token = token

		let headers = {
			'Authorization': 'Bearer ' + this.token,
			'Content-Type': 'application/json',
		},

		params = {
			sync_token: '*',
			resource_types: ["user","projects"]
		},
		
		message = "Connecting..."
		view.show(status, { message })

		fetch(todoist.url, { 
			method: 'POST',
			headers : headers,
			body: JSON.stringify(params)
		})
	
		.then(res => {
			res.json().then(data => {
				localStorage.setItem("todoist_token", todoist.token)

				view.show(settings, {
					avatar: data.user.avatar_medium || view.avatar,
					user: data.user.full_name,
					mail: data.user.email.toLowerCase(),
					tasks: view.tasks.length + " task(s)",
					title: view.pageTitle,
					preview: view.tasks[0].substring(0, 34),
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

	push(tasks = [], id = "new") {
		let project = { id, tasks },
		project_id = parseInt(project.id) || todoist.uuid(),
		commands = [],

		headers = {
			'Authorization': 'Bearer ' + this.token,
			'Content-Type': 'application/json'
		},

		message = "Proccesing..."
		view.show(status, { message })

		if (project.id == "new") {
			commands.push({
				type: "project_add",
				temp_id: project_id,
				uuid: todoist.uuid(),
				args: { name: view.pageTitle }
			})
		} 

		tasks.forEach( todo => {
			let content = view.addLinks ? `[${todo}](${view.pageUrl})` : todo
			commands.push({
				type: "item_add",
				temp_id: todoist.uuid(),
				uuid: todoist.uuid(),
				args: { content, project_id }
			})
		})

		if (commands.length == 0) {
			view.alert("No new tasks found.", 
					   "No new tasks found on this page!")

		} else {
			fetch(todoist.url, { 
				method: 'POST',
				headers : headers,
				body: JSON.stringify({ commands })
			})
			
			.then(res => {
				res.json().then(data => {
					let projectId = data.temp_id_mapping[project_id] || project_id
					window.open("https://todoist.com/showProject?id=" + projectId, "_blank")
					view.close()
				})
			})

			.catch(error => {
				view.alert("Export failed!", error);
			})
		}

	},

	uuid() {
		return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
			let d = new Date().getTime(),
			r = (d + Math.random()*16)%16 | 0
			d = Math.floor(d/16)
			return (c=='x' ? r : (r&0x7|0x8)).toString(16)
		})
	},

	async getData(project_id) {
		return await fetch("https://api.todoist.com/sync/v8/projects/get_data", { 
			method: 'POST',
			headers : {
				'Authorization': 'Bearer ' + this.token,
				'Content-Type': 'application/json'
			},
			body: JSON.stringify({ project_id })
		})
		.then( res => res.json())
		.catch(error => {
			view.alert("Todoist sync failed!", error);
		})
	},

},

tools = {
	getText(content) {
		return content.split(']')[0].replace(/\s/gi, " ").replace(/\[/gi, "")
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
					} else if (p.type == "Table") {
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
